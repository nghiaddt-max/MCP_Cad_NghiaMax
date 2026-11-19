"""Loi vao giu tuong thich voi .mcp.json cu (chay thang file nay bang python cua .venv).
Ma nguon that nam trong goi autocad_mcp/; ban cai dat cho may khac dung lenh autocad-mcp.exe."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from autocad_mcp.server import main  # noqa: E402

if __name__ == "__main__":
    main()
