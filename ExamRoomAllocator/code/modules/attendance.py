# attendance.py
import sys
from pathlib import Path
import logging
from modules.seating import process_master_excel
from modules.pdf_gen import generate_pdfs_for_output_dir
from modules.utils import setup_output_dir, create_final_zip

LOG = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


def main(argv):
    if len(argv) < 3:
        print("Usage: python attendance.py <master.xlsx> <photos_dir> [output_zip_name]")
        return 2
    master = argv[1]
    photos_dir = argv[2]
    out_name = argv[3] if len(argv) >= 4 else "Output_Zip"

    tmpdir = Path("tmp_run")
    tmpdir.mkdir(parents=True, exist_ok=True)
    output_root = setup_output_dir(str(tmpdir))

    # if master is a path, open bytes
    with open(master, "rb") as f:
        master_bytes = f.read()

    from io import BytesIO
    process_master_excel(BytesIO(master_bytes), output_root, buffer=0, filling_mode="Dense")
    generate_pdfs_for_output_dir(output_root, photos_dir, cols=3, paper_format="A4")
    zip_path = create_final_zip(output_root, out_name)
    print("Output zip:", zip_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
