
import pandas
import os
import shutil
import click

@click.command()
@click.option('-k','--key', prompt="Location of spreadsheet", help="Location of spreadsheet", type=click.Path(exists=True, readable=True, file_okay=True, dir_okay=False))
@click.option('-s', '--sourcedir', prompt="Location of input files", help="Location of input files", type=click.Path(exists=True, readable=True, file_okay=False, dir_okay=True))
@click.option('-d', '--destdir', prompt="Location for output files", help="Location of output files", type=click.Path(file_okay=False, dir_okay=True, writable=True))
@click.option('--cols', default="A,B", help="Columns in sheet containing mapping", show_default=True)
@click.option('--skiprows', default=0, help="Number of rows in sheet that need to be skipped", show_default=True)
def bulk_rename_from_excel(key="test/key.xlsx", sourcedir="test/", destdir="test/destination/", cols="F,G", skiprows=10):
    """Bulk-rename files based off of mappings contained in an Excel spreadsheet"""

    sheet = pandas.read_excel(key, usecols=cols, skiprows=skiprows, names=["code","description"])
    print(f"Reading from '{key}':")
    print("----")
    print(sheet)
    if not os.path.exists(destdir):
        print(f"Destination '{destdir}' does not exist. Creating.")
        os.mkdir(destdir)
    rows,cols = sheet.shape
    old_col = sheet.get("code")
    new_col = sheet.get("description")
    copies = []
    print("----")
    print("Planning to make the following copies:")
    print("----")
    extension = ".txt"
    for i in range(rows):
        old_prefix = str(old_col.get(i))
        old_filenames = shutil.fnmatch.filter(os.listdir(sourcedir), old_prefix+"*")
        new_prefix = str(new_col.get(i))
        for old_filename in old_filenames:
            old_filename = os.path.join(sourcedir, old_filename)
            old_extension = os.path.splitext(old_filename)[-1]
            new_filename = os.path.join(destdir, new_prefix + old_extension)
            print(f"'{old_filename}' => '{new_filename}'")
            copies.append((old_filename, new_filename))

    print("----")
    if click.confirm('Are you sure you want to continue?', default=False):
        for (old_filename, new_filename) in copies:
            shutil.copyfile(old_filename,new_filename)



if __name__ == '__main__':
    bulk_rename_from_excel()


