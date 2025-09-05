
### Site Packages in Virtual Environment
```Python
python3 -m venv --system-site-packages forensic-env
source forensic-env/bin/activate

python -c "import pyewf, pytsk3; print('Forensic libs OK')"

```