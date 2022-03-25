

# Building

Follow the OS specific installation and setup instructions below, then run `latexmk -pdf permission_CS.tex -pvc` to generate a preview of the PDF output.


# Installation and Setup 

## Windows

Install WSL2 and follow the linux instructions.

##  Linux 

### Dependencies

* [Latexmk](https://personal.psu.edu/~jcc8/software/latexmk/)
* evince PDF reader

### Installation

Install texlive by following the instructions https://www.tug.org/texlive/. 

Install latexmk using texlive utility GUI or `sudo tlmgr install latexmk` - if you're on Ubuntu, try `sudo apt-get install -y latexmk`.


### Setup

create a file in `$HOME/.latexmkrc` with the contents:

```
$dvi_previewer = 'start xdvi -watchfile 1.5';
$ps_previewer  = 'start gv --watch';
$pdf_previewer = 'start evince';
```

## Mac

### Dependencies

* [MacTex](http://www.tug.org/mactex/) - the MacOS version of tex rendering engine. 
* [Latexmk](https://personal.psu.edu/~jcc8/software/latexmk/)
* [Skim PDF reader](https://skim-app.sourceforge.io/)

### Installation

Install MacTex by downloading from the [distribution page](http://www.tug.org/mactex/) and following their instructions. 

Latexmk should come with MacText. If it doesn't come with MacTex, then you can install it with texlive utility GUI or `sudo tlmgr install latexmk` from the terminal.

Skim is a lightweight PDF reader. Install by downloading from their home page or using homebrew `brew install --cask skim`

### Setup

create a file in `$HOME/.latexmkrc` with the contents:

```
$pdf_previewer = 'open -a Skim';
$pdflatex = 'pdflatex -synctex=1 -interaction=nonstopmode';
@generated_exts = (@generated_exts, 'synctex.gz');
```

If you prefer another PDF editor, change the value for `$pdf_previewer` in that 

# Reference

* [Using Latexmk](https://mg.readthedocs.io/latexmk.html#using-latexmk)