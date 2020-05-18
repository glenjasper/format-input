format-input
======================
[![License](https://poser.pugx.org/badges/poser/license.svg)](./LICENSE)

Script that formats, in Excel files (.xlsx), the tabulated files (.csv) exported from Scopus, Web of Science, PubMed, Dimensions or a text file (.txt)

## Table of content

- [Pre-requisites](#pre-requisites)
    - [Python libraries](#python-libraries)
- [Installation](#installation)
    - [Clone](#clone)
    - [Download](#download)
- [How To Use](#how-to-use)
- [Author](#author)
- [Organization](#organization)
- [License](#license)
- [Acknowledgments](#acknowledgments)

## Pre-requisites

### Python libraries

```sh
  $ sudo apt install -y python3-pip
  $ sudo pip3 install --upgrade pip
```

```sh
  $ sudo pip3 install argparse
  $ sudo pip3 install xlsxwriter
  $ sudo pip3 install pandas
```

## Installation

### Clone

To clone and run this application, you'll need [Git](https://git-scm.com) installed on your computer. From your command line:

```bash
  # Clone this repository
  $ git clone https://github.com/glenjasper/format-input.git

  # Go into the repository
  $ cd format-input

  # Run the app
  $ python3 format-input.py --help
```

### Download

You can [download](https://github.com/glenjasper/format-input/archive/master.zip) the latest installable version of _format-input_.

## How To Use

```sh  
  $ python3 format_input.py --help
  usage: format_input.py [-h] -t {scopus,wos,pubmed,dimensions,txt} -i
                         INPUT_FILE [-o OUTPUT] [--version]

  Script que faz a formatação, em arquivos Excel (.xlsx), os arquivos tabulados
  (.csv) exportadas do Scopus, Web of Science, PubMed, Dimensions ou de um
  arquivo de texto (.txt)

  optional arguments:
    -h, --help            show this help message and exit
    -t {scopus,wos,pubmed,dimensions,txt}, --type_file {scopus,wos,pubmed,dimensions,txt}
                          scopus: Tipo de arquivo exportado do Scopus (.csv) |
                          wos: Tipo de arquivo exportado do Web of Science
                          (.csv) | pubmed: Tipo de arquivo exportado do PubMed
                          (.csv) | dimensions: Tipo de arquivo exportado do
                          Dimensions (.csv) | txt: Tipo de arquivo .txt
    -i INPUT_FILE, --input_file INPUT_FILE
                          Arquivo exportado ou de texto que contem a lista dos
                          DOIs
    -o OUTPUT, --output OUTPUT
                          Pasta de saida com a formatação nova
    --version             show program's version number and exit

  Thank you!
```

## Author

* [Glen Jasper](https://github.com/glenjasper)

## Organization
* [Molecular and Computational Biology of Fungi Laboratory](http://lbmcf.pythonanywhere.com) (LBMCF, ICB - **UFMG**, Belo Horizonte, Brazil)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details

## Acknowledgments

* Aristóteles Góes-Neto
* Rosimeire Floripes
* Joyce da Cruz Ferraz
