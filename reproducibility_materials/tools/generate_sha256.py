"""지정 폴더의 파일에 대한 SHA-256 목록을 생성한다."""
from __future__ import annotations

import argparse
import hashlib
from pathlib import Path


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("directory", help="해시를 계산할 산출물 폴더")
    parser.add_argument("--output", default="SHA256SUMS.txt")
    args = parser.parse_args()

    directory = Path(args.directory)
    files = sorted(p for p in directory.iterdir() if p.is_file() and p.name != args.output)
    lines = [f"{sha256_file(path)}  {path.name}" for path in files]
    output = directory / args.output
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Created: {output}")


if __name__ == "__main__":
    main()

