# Simple XLSX images extractor

## Setup

### ENV
```
virtualenv .env
. .env/bin/activate
pip install -r requirements.txt
python run.py -h
....
deactivate
```

### FILE

Store your file somewhere on your hard drive

Investigate it against image placement and its relation

### Launch

```
. .env/bin/activate
mkdir my_images
python run.py -i /path/to/your/file.xlsx -s MYSHEET -f 6 -r 7 -l 2 -o my_images/
deactivate
```

Enjoy!

## License

MIT LICENSE WITHOUT WARRANTY OF ANY KIND
