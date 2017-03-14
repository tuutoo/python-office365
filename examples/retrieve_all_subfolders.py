from office365api import Mail
from dotenv import load_dotenv
from os.path import join, dirname, normpath
from os import environ

from office365api.model.folder import Folder

dot_env_path = normpath(join(dirname(__file__), '../', '.env'))
load_dotenv(dot_env_path)
authorization = (environ.get('OFFICE_USER'), environ.get('OFFICE_USER_PASSWORD'))
mail = Mail(auth=authorization)


def simplest(auth):
    c = mail.folders.get_count()
    print('Folder count {0}'.format(c))

    m = mail.folders.get_all_folders()
    print('Folder names.')
    for folder in m:
        print("    {id} {name}".format(id=folder.Id, name=folder.DisplayName))
        retrieve_sub_folder(auth=auth, folder=folder, indent='    ')


def retrieve_sub_folder(auth, folder: Folder, indent):
    _indent = indent + '    '

    if folder.ChildFolderCount > 0:
        sf = mail.folders.get_sub_folders(folder.Id)
        for f in sf:
            print("{indent}{id} {name}".format(id=f.Id, name=f.DisplayName, indent=_indent))
            retrieve_sub_folder(auth=auth, folder=f, indent=_indent)


if __name__ == '__main__':
    simplest(authorization)
