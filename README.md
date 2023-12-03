# pdf2ppt
Convert PDF Slides to PowerPoint Presentations (PPT)

## Motivation
LaTeX users can easily convert the `beamer` slides in PDF to PPT.

## Dependency
- **Python >= 3.9**: This project heavily relies on [`python-pptx`](https://python-pptx.readthedocs.io/). (Other required packages: `pypdf`, `tqdm`. Check [`requirements.txt`](requirements.txt))
- [**pdf2svg**](https://github.com/dawbarton/pdf2svg): used for converting PDF to SVG
- [**Inkscape**](https://inkscape.org/): used for converting SVG to EMF

## Technical Implementation
1. The first step is to create SVG from PDF via `pdf2svg`.
2. Due to the limitation of `python-pptx`, we need to convert SVG to EMF via `inkscape`.
3. Insert EMF into PPT via `python-pptx`.

> [!NOTE]
> `python3`, `pdf2svg` and `inkscape` should be in your PATH.
> Alternatively, `--pdf2svg-path` and `inkscape-path` options can be used to specify their paths.

## Usage
### Installation
Use git to clone the repository.
```sh
git clone https://github.com/Teddy-van-Jerry/pdf2ppt.git --depth=1
```
If you only want the latest Python script, you can directly download the source file.
```sh
wget https://raw.githubusercontent.com/Teddy-van-Jerry/pdf2ppt/master/pdf2ppt
```

> [!TIP]
> Make sure you have the [dependency](#dependency) installed.

### Command Line Options
You can use `pdf2svg -h` to view all options.

### Quick Start

> [!NOTE]
> If you have not installed `pdf2ppt` to your PATH, you need to use `./pdf2ppt` in the correct directory.

**Specifying the output file name**.
```sh
pdf2ppt input.pdf output.pptx
```

**Without specifying output file name**.
The output will be `input.pptx`.
```sh
pdf2ppt input.pdf
```

**Toggle verbose mode**.
```sh
pdf2ppt input.pdf output.pptx --verbose
```

## License
Copyright ©️ 2023 Teddy van Jerry (Wuqiong Zhao).
This project is distributed under the [MIT License](LICENSE).
