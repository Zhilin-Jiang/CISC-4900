
## Dependencies

* Python v3.9
* [PyPDF4](https://github.com/claird/PyPDF4)
* argparse

## Running

From this folder, you can run the script `python3 sample.py`

# Flags

There are some customizable configurations that you can provide to this script, as viewable using the -h flag.

* Specifying directory to read (with -d or --d followed by the path)
* Specifying the output directory to save report file (with -o)
* Customize the filename for the report file

```sh
$ python3 sample.py -h
usage: sample.py [-h] [-d D] [-o O] [-f FILENAME]

Text goes here.

optional arguments:
  -h, --help            show this help message and exit
  -d D, --directory D   directory path where files are located
  -o O, --output-directory O
                        output directory
  -f FILENAME, --filename FILENAME
                        give a filename to output. otherwise it will be _report_YYYY-MM-DD.csv
```
