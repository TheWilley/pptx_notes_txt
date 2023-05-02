#######################################
# Author: TheWilley                   #
# Copyright: 2023                     #
# License: MIT                        #
# Version: 1.0.0                      #
#######################################

import click
import zipfile
import xml.etree.ElementTree as ET

@click.command()
@click.option('-i', '--input', default=None, help='The path to the pptx file.', required=True, type=str)
@click.option('-o', '--output', help='The path to the output file.', required=True, type=str)
@click.option('-p', '--prettyprint', is_flag=True, help='Write the text in a pretty format.')
@click.option('-md', '--markdown', is_flag=True, help='Write the text in a markdown format.')
@click.option('-c', '--custom', help='Write the text in a custom format using a \'.custom\' file', type=str)
def main(input, output, prettyprint, markdown, custom):
    """Script to extract the notes from a pptx file and write them to a text file."""

    # Define lists to store the contents of the notes and slides
    note_files = []
    pages = []

    # Check if input is empty
    if input is None:
        click.echo('No input file provided.')
        exit(0)

    # Check if file exists
    try:
        with open(input) as f:
            pass
    except Exception as e:
        click.echo('Error: Input file does not exist.')
        exit(0)

    # Check if input is a pptx file
    if not input.endswith('.pptx'):
        click.echo('Error: Input file is not a valid pptx file.')
        exit(0)

    # Open the pptx file
    with zipfile.ZipFile(input, 'r') as pptx:

        # Loop through all the files in the pptx file and store the contents of the notes in a list
        for file in pptx.namelist():

            # Check if the file is a notes file
            if file.startswith('ppt/notesSlides/notesSlide'):

                # Open the notes file and store the contents in list
                with pptx.open(file) as note:
                    note_files.append(note.read().decode('utf-8'))

    # Loop through the list of notes and extract the text from the tags
    for note in note_files:
        # Define lists to store the contents of the tags and the text from the tags
        all_text_tags = []
        text_from_tags = []

        # Create an ElementTree object from the note
        myroot = ET.fromstring(note)

        # Find all the note tags in the note and store them in a list
        all_text_tags.append(myroot.findall('.//{*}t'))

        # Loop through the list of tags and extract the text from the tags
        for t in all_text_tags:
            for i in t:

                # Check if its the last tag
                if i == t[-1]:
                    break

                # Append the text from the tag to the list
                text_from_tags.append(i.text)

        # Join the text from the tags and append it to the pages list
        pages.append(' '.join(text_from_tags))

    # Try to write the contents of the pages list to a file
    try:
        with open(output, 'w') as f:
            # Check if both prettyprint, markdown and custom are used
            if prettyprint and markdown or prettyprint and custom or markdown and custom:
                click.echo('Error: You can only use one of the flags [-p, --prettyprint] or [-md --markdown].')
                exit(0)

            # Mode is prettyprint
            if prettyprint:
                for page in pages:
                    f.write('Page ' + str(pages.index(page) + 1) + '\n' + '---------------\n\n')
                    f.write(page)

                    # If its the last page, don't add a new line
                    if page == pages[-1]:
                        break
                    else:
                        f.write('\n\n')

            # Mode is markdown
            elif markdown:
                for page in pages:
                    f.write('# Page ' + str(pages.index(page) + 1) + '\n')
                    f.write(page)

                    # If its the last page, don't add a new line
                    if page == pages[-1]:
                        break
                    else:
                        f.write('\n\n')

            # Mode is custom
            elif custom:
                # Check if the custom file is empty
                if custom is None:
                    click.echo('Error: No custom file provided.')
                    exit(0)

                # Check if the custom file is a .custom file
                if not custom.endswith('.custom'):
                    click.echo('Error: Not a valid \'.custom\' file.')
                    exit(0)

                # Check if the custom file exists
                try:
                    with open(custom) as f2:
                        pass
                except Exception as e:
                    click.echo('Error: \'.custom\' file does not exist.')
                    exit(0)

                # Open the custom file and store the contents in a list
                with open(custom) as f2:
                    custom = f2.read()

                for page in pages:
                    # Use regex entered from user to replace the text with the custom format
                    f.write(custom.replace("{notes}", page).replace("{slide}", str(pages.index(page) + 1)))
            
            # Mode is default
            else:
                for page in pages:
                    f.write(page)

                # If its the last page, don't add a new line
                    if page == pages[-1]:
                        break

                    f.write('\n\n')

    except Exception as e:
        click.echo("Error: {}".format(e))
        exit(0)
    
    # Print success message
    click.echo('Successfully wrote the notes to \'' + output + '\'')

if __name__ == '__main__':
    main()