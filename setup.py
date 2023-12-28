from cx_Freeze import setup, Executable

setup(
    name="NBHGTV",
    version="0.2",
    description="TinhNhuanButHGTV",
    executables=[Executable("main.py")]
)
