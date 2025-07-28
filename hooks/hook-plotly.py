import os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Substitua pelo caminho exato da pasta .dist-info do kaleido
kaleido_dist_info_path = r'C:\Users\silonei.duarte.BRUNO\AppData\Local\Programs\Python\Python313\Lib\site-packages\kaleido-0.2.1.dist-info'

kaleido_metadata = [
    (kaleido_dist_info_path, os.path.basename(kaleido_dist_info_path))
]

hiddenimports = (
    collect_submodules('plotly') +
    collect_submodules('oracledb') +
    collect_submodules('fontTools') +
    collect_submodules('cryptography') +
    collect_submodules('kaleido')
)

datas = (
    collect_data_files('plotly', include_py_files=False) +
    collect_data_files('oracledb', include_py_files=False) +
    collect_data_files('fontTools', include_py_files=False) +
    collect_data_files('cryptography', include_py_files=False) +
    collect_data_files('kaleido', include_py_files=False) +
    kaleido_metadata
)
