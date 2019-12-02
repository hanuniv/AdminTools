
"""
This is the script that generate a beamer file from tex file
by simply make each theorem, lemma, etc.. a frame.
The script guess a frame title from theorem name or label


usage:

python create-beamer [input_tex_file] [output_tex_file]

Default: 
    input_tex_file = main.tex
    output_tex_file = main-beamer.tex

One has the option to confirm the name during execution

assumptions:
1. "\input" occupy one line
2. "\begin" and "\end" occupy one line, cannot be used on the same line
3. no nested environment of the same name
"""

import re
import time
import sys
import os.path


def lines(filename):
    """
    A generator that deal with input file line-wise, 
    It handles \input command
    """
    input_pattern = re.compile(r"\\input{(.*?)}")

    with open(filename, 'r') as file:
        for line in file:
            # if is input command, open the file
            input_match = input_pattern.search(line)
            if input_match:
                input_filename = input_match.group(1)
                if not input_filename.endswith('.tex'):
                    input_filename += ".tex"
                yield from lines(input_filename)
                yield '\n'   # add an empty line to separate input files
            else:
                yield line


def blocks(filename, capture_names):
    """
    A generator that combine parts of file into blocks
    capture_names is a list of names to be captured,
    such as equation, thm, etc. 
    The blocks within a captured blocks will not be output separately

    Yield a tuple of two strings: 
        block, captured_block_name
    """
    begin_patterns = {name: re.compile(
        r"\\begin{{{name}}}".format(name=name)) for name in capture_names}

    end_patterns = {name: re.compile(
        r"\\end{{{name}}}".format(name=name)) for name in capture_names}

    block = []
    begin_match = None  # mark if has begin a block
    block_name = None   # remember block name that is captured

    for line in lines(filename):

        if begin_match:
            # if has started a block
            block.append(line)
            if end_patterns[block_name].search(line):
                # if a block ended, restart capture
                yield "".join(block), block_name
                block = []
                begin_match = None
                block_name = None
        else:
            # outside a block, check if a block should be started
            # assumption: "\begin" occupy one line
            for block_name in capture_names:
                begin_match = begin_patterns[block_name].search(line)
                if begin_match:
                    block.append(line)
                    break
            else:
                # no block is started, yield this line
                yield line, None


def checkfilename(infilename, outfilename):
    """
    Input: default infilename, default outfilename
    Return a tuple: infilename, outfilename
    the function make sure input filename exists, outfilename != infilename
    and confirms overwriting if outfilename exists
    """
    # check filenames
    while True:
        # make sure input file exists
        while not os.path.isfile(infilename):
            print("Input filename '{}' does not exist".format(infilename))
            infilename = input('Please enter input file name:')

        # make sure output != input and confirm overwrite
        while os.path.exists(outfilename):
            if 'n' != input("Output filename '{}' exists.\n Overwrite? [y]/n >".format(outfilename)):
                if outfilename != infilename:
                    break
                else:
                    print("Same name for input and output file")
                    outfilename = input('Please enter output file name:')
            else:
                outfilename = input('Please enter output file name:')

        # Summary and confirmation
        print('\nInfile = ' + infilename)
        print('Outfile = ' + outfilename, end='')
        if os.path.exists(outfilename):
            print(" [**Overwrite**]")
        else:
            print()
        if 'n' != input('Confirm? [y]/n >'):
            break

    return infilename, outfilename


if __name__ == '__main__':

    DEBUG = False

    if DEBUG:
        infilename, outfilename = 'main.tex', 'main-beamer.tex'  # ! Debug setting
    else:
        print("Welcome to Beamer generator by Fh.\n")
    
        # default option:
        infilename, outfilename = 'main.tex', 'main-beamer.tex'  # ! Debug setting
        
        # checks availability of commandline arguments, overwrite defaults
        if len(sys.argv) > 1:
            infilename = sys.argv[1]
        if len(sys.argv) > 2:
            outfilename = sys.argv[2]

        infilename, outfilename = checkfilename(infilename, outfilename)

    frametitle_pat1 = re.compile(r'\\begin\{.*?\}\[(.*?)\]')
    frametitle_pat2 = re.compile(r'\\label\{(?:.*?\:)?(.*?)\}')
    with open(outfilename, 'w') as fout:
        for block, name in blocks(infilename, ["theorem", "thm", "lemma", "prop", "cor", "eg", "conj", "remark", "assume", "definition", "figure", "itemize"]):
            if name is not None:
                fout.write(r"\begin{frame}" + '\n')
                # guess frame title first from theorem name
                m = frametitle_pat1.match(block.strip())
                if m:
                    title = m.group(1)
                else:  # else from label
                    m = frametitle_pat2.search(block)
                    if m:
                        title = m.group(1)
                    else:
                        title = " "
                fout.write(r"\frametitle{{{0}}}".format(title) + "\n")
                fout.write(block)
                fout.write(r"\end{frame}" + '\n\n')

    print('Done. Have a nice day.')

    if not DEBUG:
        time.sleep(1)
