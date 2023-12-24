import click
import os
from collections import Counter
from PIL import Image
import pyfiglet
from colorama import Fore
from pdf2image import convert_from_path
import pdfplumber
from docx import Document

# Functie voor het bepalen van het outputformaat
def get_output_format(extension):
    format_options = {
        '.png': ['jpg', 'pdf'],
        '.jpg': ['png', 'pdf'],
        '.jpeg': ['png', 'pdf'],
        # Andere ondersteunde afbeeldingsformaten
    }
    if extension in format_options:
        click.echo(f"Beschikbare formaten voor conversie van {extension}: {', '.join(format_options[extension])}")
        return click.prompt("Kies een outputformaat", type=click.Choice(format_options[extension]))
    else:
        return None

# Functie voor het converteren van afbeeldingen
def convert_image(input_path, output_path, output_format):
    with Image.open(input_path) as img:
        img.save(output_path, format=output_format)

# Functie voor het tellen van bestanden op extensie
def get_file_count_by_extension(directory):
    extensions = []
    for filename in os.listdir(directory):
        if os.path.isfile(os.path.join(directory, filename)):
            ext = os.path.splitext(filename)[1].lower()
            extensions.append(ext)
    return Counter(extensions)

# Functie om PDF naar afbeelding te converteren
def convert_pdf_to_image(pdf_path, output_path, image_format='png'):
    images = convert_from_path(pdf_path)
    for i, image in enumerate(images):
        image_file = f'{output_path}_{i}.{image_format}'
        image.save(image_file, image_format.upper())

# Functie om PDF naar tekst te converteren
def convert_pdf_to_text(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
        return text


# Functie om tekst op te slaan in een Word-document
def save_text_to_word(text, output_path):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(output_path)

# Functie om het afbeeldingsformaat voor PDF-conversie te kiezen
def get_image_format():
    formats = ['png', 'jpg']
    chosen_format = click.prompt("Kies een afbeeldingsformaat voor de PDF-conversie", type=click.Choice(formats))
    return chosen_format

@click.command()
def main():
    click.echo(Fore.GREEN + pyfiglet.figlet_format("QuickConv") + Fore.RESET)
    click.echo(Fore.YELLOW + "Welkom bij quickconv. Geef de map op waar de bestanden die je wilt converteren zich bevinden:" + Fore.RESET)
    directory = click.prompt(Fore.BLUE + "Voer het pad in" + Fore.RESET)

    file_count = get_file_count_by_extension(directory)
    click.echo(Fore.CYAN + f"Quickconv heeft {sum(file_count.values())} bestanden gevonden:" + Fore.RESET)
    for ext, count in file_count.items():
        click.echo(Fore.CYAN + f"{count} {ext[1:]}" + Fore.RESET)

    ext_to_convert = click.prompt("Welke bestandsformatie wil je converteren?").lower()
    output_format = get_output_format('.' + ext_to_convert)
    if output_format is None and ext_to_convert != 'pdf':
        click.echo("Geen geldige conversieopties beschikbaar. Operatie gestopt.")
        return

    all_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith(ext_to_convert)]
    if click.confirm("Wil je alle bestanden converteren?"):
        selected_files = all_files
    else:
        for i, file in enumerate(all_files, start=1):
            click.echo(f"{i}. {file}")
        file_indices = click.prompt("Voer de nummers in van de bestanden die je wilt converteren", type=str)
        indices = [int(i) - 1 for i in file_indices.split(",")]
        selected_files = [all_files[i] for i in indices]

    output_dir = click.prompt("Voer het pad in voor de output (standaard is huidige map)", default=directory)

    if ext_to_convert == 'pdf':
        for pdf_file in selected_files:
            if click.confirm("Wil je de PDF converteren naar een afbeelding?"):
                image_format = get_image_format()
                output_file_base = os.path.splitext(os.path.basename(pdf_file))[0]
                output_file_path = os.path.join(output_dir, output_file_base)
                convert_pdf_to_image(pdf_file, output_file_path, image_format)
            elif click.confirm("Wil je de PDF converteren naar tekst/Word?"):
                text = convert_pdf_to_text(pdf_file)
                word_output_path = os.path.join(output_dir, os.path.splitext(os.path.basename(pdf_file))[0] + '.docx')
                save_text_to_word(text, word_output_path)
    else:
        for input_path in selected_files:
            filename = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(output_dir, f"{filename}.{output_format}")
            convert_image(input_path, output_path, output_format)
            click.echo(f'Converted {input_path} to {output_path}')

    click.echo(Fore.GREEN + "Conversie voltooid." + Fore.RESET)

if __name__ == '__main__':
    main()
