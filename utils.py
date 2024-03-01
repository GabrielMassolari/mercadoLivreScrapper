from definitions import ROOT_DIR
import os


def get_complete_root_file_path(filename):
    return os.path.join(ROOT_DIR, filename)
