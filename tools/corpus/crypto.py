from __future__ import annotations

import os
from pathlib import Path


def get_fernet_key_from_env(env_var: str = "CORPUS_ENCRYPTION_KEY") -> str:
    key = os.environ.get(env_var)
    if not key:
        raise RuntimeError(
            f"Missing {env_var}. Generate one with `python tools/corpus/keygen.py`."
        )
    return key


def encrypt_file(input_path: Path, output_path: Path, *, fernet_key: str) -> None:
    from cryptography.fernet import Fernet

    data = input_path.read_bytes()
    f = Fernet(fernet_key.encode("utf-8"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(f.encrypt(data))


def decrypt_file(input_path: Path, output_path: Path, *, fernet_key: str) -> None:
    from cryptography.fernet import Fernet

    data = input_path.read_bytes()
    f = Fernet(fernet_key.encode("utf-8"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(f.decrypt(data))

