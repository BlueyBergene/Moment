from pprint import pprint
import pandas as pd
import pathlib
import shutil
import click
import sys


# TODO : Add GUI, I think I'd like to use PySimpleGUI for this.
# TODO : Make the script change working directory to the dir where the files are to be moved FROM

def get_files(excel_file: str) -> list:
    """Extracts filepaths from the given Excel document.

    Parameters:
        excel_file (str): The path to the Excel file.

    Returns:
        list: Lists of paths"""


    file = pathlib.Path(excel_file)
    if file.is_file():
        click.echo(click.style("Success", fg="green") + " - " + "Source file found.")
        # TODO : I need to restructure how my script read the excel file.
        #  Column A should be 'document name',
        #  Column B should be 'Source',
        #  Column C should be 'Destination' - Currently the destination is passed as an option when launching the script, and it's used for ALL files..
        #  Currently my script just grabs whatever values are in the Excel sheet and assumes they are paths..
        dataframe2 = pd.read_excel(excel_file, header=None, index_col=None)
        files_lists_in_list = dataframe2.values.tolist()
        files_list = merge_lists(files_lists_in_list)
        click.echo(click.style("Success", fg="green") + " - " + "Source file read.")
        return files_list
    elif file.is_dir():
        # TODO : I could add folder handling. When a folder is specified it enumerates the files in that dir instead of getting files form Excel.
        click.echo(click.style("Error", fg="red") + " - " + "Given path is a folder, not an Excel file." + click.style(" Aborting!", fg="red"))
    else:
        click.echo(click.style("Error", fg="red") + " - " + "Source file not found." + click.style(" Aborting!", fg="red"))
    sys.exit()




def merge_lists(files_list: list):
    """Merges lists that contain lists into JUST a list.

    Parameters:
        files_list (list): List of lists.

    Returns:
        list: Merged list containing only paths."""
    merged_list = []
    for list in files_list:
        merged_list.append(list[0])
    return merged_list

def sort_paths(files_list: list) -> dict:
    """Sorts files under their parent folders.

    Parameters:
        files_list (list): List of absolute file paths.

    Returns:
       Dictionary, the files are put into lists which in turn is store in the dict under their respective parents."""
    files_dict: dict = {}
    for file in files_list:
        file_path = pathlib.Path(file)
        absolute_path = file_path.resolve()
        parent = str(absolute_path.parent)
        file_name = absolute_path.name
        if not parent in files_dict:
            files_dict[parent] = []
        files_dict[parent].append(file_name)
    return files_dict

def enum_files(destination: str, files_dict: dict, move=False, test=False):
    skipped_files = []
    dest = pathlib.Path(destination)
    if not dest.exists():
        click.echo(click.style("Error", fg="red") + " - " + "Given destination is not a valid folder." + click.style(" Aborting!", fg="red"))
        sys.exit()
    else:
        click.echo("\n" + click.style("Success", fg="green") + " - " + "Destination exists.", color=True)

    for source, files_list in files_dict.items():
        for file in files_list:
            full_source = pathlib.Path(source, file)
            full_destination = pathlib.Path(dest, file)
            if full_source.is_file():
                if not move and not test:
                    #shutil.copy2(full_source, full_destination)
                    shutil.copyfile(full_source, full_destination)
                elif move and not test:
                    shutil.move(full_source, full_destination)

                click.echo(click.style("Success", fg="green") + " - " + f"{'Copied' if not move else 'Moved'} file ({click.style(file, fg='yellow')}) from '{click.style(source, fg='magenta')}' to '{click.style(destination, fg='cyan')}'")
            else:
                # TODO : If a file cannot be moved for any reason.. Add it to a list to display at the end of the script.
                click.echo(click.style("Warning", fg="yellow") + " - " + f"Skipping - Cannot find file: {full_source}")
                skipped_files.append(source + "\\" + file)
    return skipped_files


@click.command()
@click.option("-s", "--source", help="Source excel file. If the path contains spaces, please surround them with quotes.", required=True)
@click.option("-d", "--dest", help="Destination folder. If the path contains spaces, please surround them with quotes.", required=True)
@click.option("-m", "--move", is_flag=True, default=False, help="Set this flag to move the files instead of copying.")
@click.option("-t", '--test', is_flag=True, default=False, help="Set this flag for a test run.")
def main(source, dest, move, test):
    files = get_files(source)
    files = sort_paths(files)
    skipped_files = enum_files(dest, files, move=move, test=test)
    pprint(skipped_files)

if __name__ == "__main__":
    main()
