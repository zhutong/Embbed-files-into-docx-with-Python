# embeddocx-01.py
import os
import shutil
import zipfile

docx_fn = 'demo.docx'
extract_folder = 'extrated'
this_path = os.path.dirname(os.path.abspath(__file__))
src_docx_fn = os.path.join(this_path, docx_fn)


def unzip_docx():
    shutil.rmtree(extract_folder, ignore_errors=True)
    os.mkdir(extract_folder)
    os.chdir(extract_folder)
    with zipfile.ZipFile(src_docx_fn) as azip:
        azip.extractall()


if __name__ == '__main__':
    unzip_docx()
