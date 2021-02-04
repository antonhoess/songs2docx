# Songs2Docx
Converts TXT files in a INI-like, but proprietary format (used for simplicity) to a DOCX file in a certain format using python-docx.

## Install
### Create a conda environment
```bash
conda init bash # => Open new terminal
conda create --name songs2docx python=3.7
conda install --name songs2docx python-docx
```

## Run
### Activate the conda environment and start the program
```bash
cd songs2docx/
conda activate songs2docx
python songs2docx.py songs/*.txt --output=output
```
