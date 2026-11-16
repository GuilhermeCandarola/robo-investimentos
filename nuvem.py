# -*- coding: utf-8 -*-
"""Roda o robo de investimentos num servidor, usando o Google Drive como pasta de arquivos.

Cada execucao:
  1. baixa do Drive o que o robo le (Minha_Carteira.xlsx e Plano_de_Acao.xlsx, onde ficam seus parametros);
  2. roda o robo normalmente (robo.py) numa pasta local de trabalho;
  3. envia de volta ao Drive os arquivos que mudaram (Painel, Plano, Dashboard.html, Dashboard.pdf e,
     se o robo ajustou, Minha_Carteira.xlsx).
Se voce editou um arquivo no Drive enquanto o robo trabalhava, o arquivo do Drive NAO e sobrescrito
(o robo tenta de novo na proxima rodada).

Configuracao em config.ini:
  [nuvem]
  ativo = sim
  pasta_drive_id = <ID da pasta Investimentos no Drive>
  credenciais = /opt/robo/segredos/conta_servico.json
  pasta_trabalho = /opt/robo/drive
"""
import configparser
import hashlib
import io
import logging
import os
import subprocess
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent
log = logging.getLogger("robo.nuvem")

# o que o robo LE do Drive (baixa antes de rodar)
ENTRADAS = ["Minha_Carteira.xlsx", "Plano_de_Acao.xlsx"]
# o que o robo GERA (envia depois). Minha_Carteira so se o robo a alterou.
SAIDAS = ["Painel_Investimentos.xlsx", "Plano_de_Acao.xlsx", "Dashboard.html", "Dashboard.pdf", "Minha_Carteira.xlsx"]


def md5(p):
    h = hashlib.md5()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


# ------------------------------------------------------------------ Google Drive
class Drive:
    """Acesso a UMA pasta do Drive por conta de servico (nada alem dessa pasta e visivel)."""

    def __init__(self, pasta_id, credenciais):
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        cred = service_account.Credentials.from_service_account_file(
            credenciais, scopes=["https://www.googleapis.com/auth/drive"])
        self.svc = build("drive", "v3", credentials=cred, cache_discovery=False)
        self.pasta = pasta_id

    def info(self, nome):
        q = f"'{self.pasta}' in parents and name = '{nome}' and trashed = false"
        r = self.svc.files().list(q=q, fields="files(id,name,modifiedTime,md5Checksum)",
                                  orderBy="modifiedTime desc", supportsAllDrives=True,
                                  includeItemsFromAllDrives=True).execute()
        f = r.get("files", [])
        return f[0] if f else None

    def baixar(self, info, destino):
        from googleapiclient.http import MediaIoBaseDownload
        req = self.svc.files().get_media(fileId=info["id"], supportsAllDrives=True)
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, req)
        feito = False
        while not feito:
            _, feito = dl.next_chunk()
        tmp = destino.with_name("~dl_" + destino.name)
        tmp.write_bytes(buf.getvalue())
        os.replace(tmp, destino)

    def enviar(self, info, origem):
        from googleapiclient.http import MediaFileUpload
        media = MediaFileUpload(str(origem), resumable=False)
        self.svc.files().update(fileId=info["id"], media_body=media, supportsAllDrives=True).execute()


class DriveLocal:
    """Drive de mentira (uma pasta comum) para testar o fluxo sem internet."""

    def __init__(self, pasta):
        self.p = Path(pasta)

    def info(self, nome):
        f = self.p / nome
        if not f.exists():
            return None
        return {"id": nome, "name": nome, "modifiedTime": f"{f.stat().st_mtime_ns}", "md5Checksum": md5(f)}

    def baixar(self, info, destino):
        destino.write_bytes((self.p / info["id"]).read_bytes())

    def enviar(self, info, origem):
        (self.p / info["id"]).write_bytes(Path(origem).read_bytes())


# ------------------------------------------------------------------ fluxo
def rodar(drive, trabalho, py=sys.executable, timeout=900):
    trabalho = Path(trabalho)
    trabalho.mkdir(parents=True, exist_ok=True)
    baixados = {}
    for nome in ENTRADAS:
        inf = drive.info(nome)
        if not inf:
            log.warning("%s nao encontrado na pasta do Drive; o robo cria um modelo", nome)
            continue
        loc = trabalho / nome
        if not loc.exists() or md5(loc) != inf.get("md5Checksum"):
            drive.baixar(inf, loc)
        baixados[nome] = (inf["modifiedTime"], md5(loc))
    log.info("Entradas baixadas: %s", ", ".join(baixados) or "nenhuma")

    env = dict(os.environ, ROBO_PASTA_SAIDA=str(trabalho))
    r = subprocess.run([py, str(BASE / "robo.py")], cwd=str(BASE), capture_output=True, text=True, timeout=timeout, env=env)
    for linha in (r.stdout or "").splitlines():   # diagnostico sem dados financeiros
        if linha.startswith("[dl_ebitda]"):
            print(linha, flush=True)
    if r.returncode != 0:
        detalhe = "" if os.environ.get("GITHUB_ACTIONS") else ": " + (r.stderr or r.stdout)[-400:]
        log.error("robo.py terminou com erro %s%s", r.returncode, detalhe)

    enviados, pulados = [], []
    for nome in SAIDAS:
        loc = trabalho / nome
        if not loc.exists():
            continue
        h = md5(loc)
        inf = drive.info(nome)
        if inf is None:
            log.warning("%s nao existe no Drive; crie/copie um arquivo com esse nome na pasta (a conta de servico so atualiza)", nome)
            pulados.append(nome)
            continue
        if h == inf.get("md5Checksum"):
            continue
        if nome in baixados:
            # arquivo que voce tambem edita: so envia se o robo mudou E ninguem mexeu no Drive enquanto isso
            if h == baixados[nome][1]:
                continue
            if inf["modifiedTime"] != baixados[nome][0]:
                log.warning("%s foi alterado no Drive durante a rodada; nao sobrescrevo (proxima rodada refaz)", nome)
                pulados.append(nome)
                continue
        drive.enviar(inf, loc)
        enviados.append(nome)
    try:
        inf = drive.info("robo_log_nuvem.txt")
        if inf:
            linhas = (BASE / "robo.log").read_text(encoding="utf-8", errors="replace").splitlines()[-150:]
            tmp = trabalho / "robo_log_nuvem.txt"
            tmp.write_text("\n".join(linhas) + "\n", encoding="utf-8")
            drive.enviar(inf, tmp)
    except Exception as e:  # noqa
        log.warning("Log nao enviado ao Drive: %s", e)
    log.info("Enviados ao Drive: %s%s", ", ".join(enviados) or "nada novo",
             f" | pulados: {', '.join(pulados)}" if pulados else "")
    return enviados, pulados


def main():
    ci = bool(os.environ.get("GITHUB_ACTIONS"))
    console = logging.StreamHandler()
    console.setLevel(logging.ERROR if ci else logging.INFO)  # no GitHub (publico) nada financeiro vai ao console
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.FileHandler(BASE / "robo.log", encoding="utf-8"), console])
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read(BASE / "config.ini", encoding="utf-8-sig")
    if cfg.get("nuvem", "ativo", fallback="nao").strip().lower() not in ("sim", "s", "1", "true"):
        log.error("[nuvem] ativo = sim nao configurado em config.ini")
        return 2
    trabalho = cfg.get("nuvem", "pasta_trabalho", fallback=str(BASE / "drive"))
    falso = None
    if "--drive-falso" in sys.argv:
        falso = sys.argv[sys.argv.index("--drive-falso") + 1]
    # trava para nao rodar duas vezes ao mesmo tempo
    lock = BASE / "cache" / "nuvem.lock"
    lock.parent.mkdir(exist_ok=True)
    if lock.exists() and time.time() - lock.stat().st_mtime < 1500:
        log.info("Rodada anterior ainda em andamento; pulando")
        return 0
    lock.write_text(str(os.getpid()))
    try:
        drive = DriveLocal(falso) if falso else Drive(cfg.get("nuvem", "pasta_drive_id"), cfg.get("nuvem", "credenciais"))
        rodar(drive, trabalho)
    finally:
        try:
            lock.unlink()
        except OSError:
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
