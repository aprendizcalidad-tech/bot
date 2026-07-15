"""Convierte el JSON de una cuenta de servicio en un bloque TOML seguro.

Uso:
    python preparar_secreto_drive.py ruta/al/archivo.json
"""

from __future__ import annotations

import json
import sys
from pathlib import Path


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit(
            "Uso: python preparar_secreto_drive.py ruta/al/archivo.json"
        )

    path = Path(sys.argv[1])
    if not path.exists():
        raise SystemExit(f"No existe el archivo: {path}")

    data = json.loads(path.read_text(encoding="utf-8"))
    required = {"type", "client_email", "private_key", "token_uri"}
    missing = sorted(field for field in required if not data.get(field))
    if missing:
        raise SystemExit("Faltan campos: " + ", ".join(missing))

    compact = json.dumps(data, ensure_ascii=False, indent=2)
    print("DRIVE_SERVICE_ACCOUNT_JSON = '''")
    print(compact)
    print("'''")


if __name__ == "__main__":
    main()
