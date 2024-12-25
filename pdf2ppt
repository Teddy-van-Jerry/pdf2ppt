#!/usr/bin/env python3

"""pdf2ppt

Convert PDF Slides to PowerPoint Presentations (PPT)
Author: Teddy van Jerry (Wuqiong Zhao)
License: MIT
GitHub: https://github.com/Teddy-van-Jerry/pdf2ppt
"""

import argparse
import os
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt
from pypdf import PdfReader
import shutil
from tqdm import tqdm

TMP_DIR_NAME = '_pdf2ppt.tmp'
PDF2PPT_VERSION = '1.0'
ERR_PDF2SVG = 101
ERR_SVG2EMF = 102

def pdf2svg(pdf_path: Path, pdf2svg_path: Path='pdf2svg', verbose=False) -> bool:
    """Convert PDF to SVG using pdf2svg

    Args:
        pdf_path (Path): Path to the input PDF
        pdf2svg_path (Path): Path to pdf2svg executable
        verbose (bool): Verbose mode
    """
    # create tmp dir, remove contents if exists
    # the dir is at the level of the input file
    tmp_dir = pdf_path.parent / TMP_DIR_NAME
    tmp_dir.mkdir(parents=True, exist_ok=True)
    pdf_name = pdf_path.stem
    # run shell script, convert pdf to svg
    pdf2svg_cmd = f'{pdf2svg_path} {pdf_path} {tmp_dir / pdf_name}_%d.svg all'
    if verbose:
        print(f'pdf2svg command: {pdf2svg_cmd}')
    return os.system(pdf2svg_cmd) == 0

def svg2emf(pdf_reader: PdfReader, pdf_path: Path, inkscape_path: Path='inkscape', verbose=False, no_check=False) -> [bool, list]:
    """Convert SVG to EMF using inkscape

    Args:
        pdf_reader (PdfReader): PdfReader of the input PDF
        pdf_path (Path): Path to the input PDF
        pdf2svg_path (Path): Path to pdf2svg executable
        verbose (bool): Verbose mode
        no_check (bool): Do not check SVG filters
    """
    tmp_dir = pdf_path.parent / TMP_DIR_NAME
    pdf_name = pdf_path.stem
    # run shell script, convert pdf to svg
    svg2emf_cmd_template = f'{inkscape_path} --export-type="emf" {tmp_dir / pdf_name}_%d.svg'
    pdf_pages = len(pdf_reader.pages)
    pages_with_filters = []
    pages_with_filters_svg_dir = pdf_path.parent / f'{pdf_name}_svg'
    if verbose:
        print(f'svg2emf command: {svg2emf_cmd_template}')
        print(f'Processing {pdf_pages} pages from svg to emf...')
    for i in tqdm(range(0, pdf_pages), desc='svg2emf', leave=verbose):
        svg_path = tmp_dir / f'{pdf_name}_{i + 1}.svg'
        svg2emf_cmd = f'{inkscape_path} --export-type="emf" {svg_path}'
        if os.system(svg2emf_cmd) != 0:
            print(f'Error while running \'{svg2emf_cmd}\'')
            return [False, []]
        if not no_check:
            # open SVG file and check if there is string "filter=" within it.
            svg_file = open(svg_path, 'r')
            svg_content = svg_file.read()
            if 'filter=' in svg_content:
                pages_with_filters.append(i + 1)
                # create the directory if first time
                if len(pages_with_filters) == 1:
                    # remove the contents within first
                    shutil.rmtree(pages_with_filters_svg_dir, ignore_errors=True)
                    pages_with_filters_svg_dir.mkdir(parents=True, exist_ok=True)
                # copy the SVG file to the directory
                shutil.copy(svg_path, pages_with_filters_svg_dir / f'{pdf_name}_{i + 1}.svg')
    return [True, pages_with_filters]

def emf2ppt(pdf_reader: PdfReader, pdf_path: Path, ppt_path: Path, verbose=False):
    tmp_dir = pdf_path.parent / TMP_DIR_NAME
    pdf_name = pdf_path.stem
    slide_pages = len(pdf_reader.pages)

    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    # PPT does not support different slide size within one file, so we use the first page as reference.
    pdf_box = pdf_reader.pages[0].mediabox
    slide_width, slide_height = pdf_box[2], pdf_box[3]
    prs.slide_width = Pt(slide_width)
    prs.slide_height = Pt(slide_height)

    # Add metadata.
    slide_title = pdf_reader.metadata.title
    slide_author = pdf_reader.metadata.author
    slide_creator = pdf_reader.metadata.creator
    slide_producer = pdf_reader.metadata.producer
    slide_subject = pdf_reader.metadata.subject
    slide_created = pdf_reader.metadata.creation_date

    if slide_title is not None:
        prs.core_properties.title = slide_title
    if slide_author is not None:
        prs.core_properties.author = slide_author
    if slide_subject is not None:
        prs.core_properties.subject = slide_subject
    if slide_created is not None:
        prs.core_properties.created = slide_created
    prs.core_properties.comments = 'Generated using pdf2ppt (https://github.com/Teddy-van-Jerry/pdf2ppt).'
    if slide_producer is not None:
        prs.core_properties.comments += '\nPDF Producer: ' + slide_producer
    if slide_creator is not None:
        prs.core_properties.comments += '\nPDF Creator: ' + slide_creator
    prs.core_properties.last_modified_by = 'pdf2ppt'

    # Embed all emf images to ppt.
    if verbose:
        print(f'Processing {slide_pages} pages from emf to ppt...')
    for i in tqdm(range(0, slide_pages), desc='emf2ppt', leave=verbose):
        slide = prs.slides.add_slide(blank_slide_layout)
        emf_path = f'{tmp_dir / pdf_name}_{i + 1}.emf'
        slide.shapes.add_picture(emf_path, Pt(0), Pt(0), width=prs.slide_width)
    
    prs.save(ppt_path)

def clean_tmp(pdf_path: Path, verbose=False) -> bool:
    tmp_dir = pdf_path.parent / TMP_DIR_NAME
    shutil.rmtree(tmp_dir)
    if verbose:
        print('Finished cleaning up temporary files.')

def main():
    # command line options
    parser = argparse.ArgumentParser(prog='pdf2ppt', description='Convert PDF Slides to PowerPoint Presentations (PPT)')
    parser.add_argument('-v', '--version', action='version', version=f'%(prog)s {PDF2PPT_VERSION}')
    parser.add_argument('input', metavar='input', type=Path, help='PDF file input')
    parser.add_argument('output', metavar='output', type=Path, help='PPT file output (default as ext replacement)', nargs='?')
    parser.add_argument('--verbose', action='store_true', help='Verbose mode')
    parser.add_argument('--no-clean', action='store_true', help='Do not clean temporary files')
    parser.add_argument('--no-check', action='store_true', help='Do not check SVG filters')
    parser.add_argument('--pdf2svg-path', metavar='pdf2svg_path', type=Path, help='Path to pdf2svg (default: pdf2svg)', default='pdf2svg')
    parser.add_argument('--inkscape-path', metavar='inkscape_path', type=Path, help='Path to inkscape (default: inkscape)', default='inkscape')
    args = parser.parse_args()

    if (args.verbose):
        print(f"Processing PDF file: {args.input}")
    if not pdf2svg(args.input, args.pdf2svg_path, args.verbose):
        print('ERROR: Failed to run pdf2svg!')
        exit(ERR_PDF2SVG)
    pdf_reader = PdfReader(args.input)
    svg2emf_ok, pages_with_filters = svg2emf(pdf_reader, args.input, args.inkscape_path, args.verbose, args.no_check)
    if not svg2emf_ok:
        print('ERROR: Failed to run svg2emf!')
        exit(ERR_SVG2EMF)
    else:
        if not args.no_check and len(pages_with_filters) > 0:
            print(f'WARNING: Pages {pages_with_filters} may not be correct, please double check.'
                   '\n         You can manually copy the generated SVG images to PPT.'
                   '\n         (More info: https://github.com/Teddy-van-Jerry/pdf2ppt/issues/1)')

    ppt_path = args.input.with_suffix('.pptx') if args.output is None else args.output
    emf2ppt(pdf_reader, args.input, ppt_path, args.verbose)
    if not args.no_clean:
        clean_tmp(args.input)
    if args.verbose:
        print(f'Successfully generate PPT file: {ppt_path}')

if __name__ == '__main__':
    main()
