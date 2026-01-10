#!/usr/bin/env python3

from __future__ import annotations


def main() -> None:
    from cryptography.fernet import Fernet

    print(Fernet.generate_key().decode("utf-8"))


if __name__ == "__main__":
    main()

