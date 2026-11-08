"""Konteks aplikasi: secrets, config, folder kerja sementara."""
import os
import tempfile
from dataclasses import dataclass

from . import auth
from .config import build_config


@dataclass
class Ctx:
    secrets: object
    cfg: dict
    workdir: str
    config_path: str

    @property
    def email(self):
        return auth.service_account_email(self.secrets)

    def services(self):
        return auth.build_services(self.secrets)


def make_ctx(secrets, workdir=None):
    cfg_over = auth._plain(secrets["config"]) if "config" in secrets else {}
    cfg = build_config(cfg_over)
    workdir = workdir or tempfile.mkdtemp(prefix="nivea_hub_")
    path = auth.write_runtime_files(workdir, secrets, cfg)
    return Ctx(secrets=secrets, cfg=cfg, workdir=workdir, config_path=path)
