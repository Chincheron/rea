import util.file_util as file_util

#Project Constants
REPO_DIR = file_util.find_repository_root()
PROJECT_BASE_DIR = (REPO_DIR / 'projects' / 'report')
CONFIG_DIR = PROJECT_BASE_DIR / 'config'
RESULTS_DIR = (PROJECT_BASE_DIR / 'results')