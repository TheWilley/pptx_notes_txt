# pptx_notes_txt
A simple python script to extract notes from a powerpoint presentation and save them to a text file.

## Requirements
* Python 3.6+
* Click

## Installation
```
pip install -r requirements.txt
```

## Usage

```
Usage: pptx_notes_txt.py [OPTIONS]

  Script to extract the notes from a pptx file and write them to a text file.

Options:
  -i, --input TEXT   The path to the pptx file.  [required]
  -o, --output TEXT  The path to the output file.  [required]
  -p, --prettyprint  Write the text in a pretty format.
  -md, --markdown    Write the text in a markdown format.
  -c, --custom TEXT  The path to a custom format file.
  --help             Show this message and exit.
python pptx_notes_txt.py -i <inputfile> -o <outputfile>
```

### Try it out
Go to the root directory of the project and run the following command:
```
python pptx_notes_txt.py -i ./test/Presentation.pptx -o ./test/notes.txt -p
```

### Formats
If you want to format the text, you can use the `--prettyprint` option to write the notes in a pretty format or the `--markdown` option to write the notes in a markdown format. 

A custom format can also be specified by creating a file with the `.custom` extension and passing it to the `--custom` option. The file should contain two placeholders: `{slide}` and `{notes}`. The `{slide}` placeholder will be replaced with the slide number and the `{notes}` placeholder will be replaced with the notes. For example, the following custom format will write the notes in a markdown format:
```
# Slide {slide}
{notes}


```

You can try it out with the following command:
```
python pptx_notes_txt.py -i ./test/Presentation.pptx -o ./test/notes.txt -c ./test/myformat.custom
```

Which yeilds the following output:
```markdown
# Slide 1
This is notes in the first slide

# Slide 2
Here are some more notes...

# Slide 3
…and even more here!
```

## License
MIT License