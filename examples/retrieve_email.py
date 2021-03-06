from office365api import Mail
from dotenv import load_dotenv
from os.path import join, dirname, normpath
from os import environ

dot_env_path = normpath(join(dirname(__file__), '../', '.env'))
load_dotenv(dot_env_path)


def simplest(auth):
    mail = Mail(auth=auth)
    m = mail.inbox.get_messages()
    print('simplest {count}'.format(count=(len(m))))
    for message in m:
        print(message.Subject)


def simple_by_folder(auth):
    mail = Mail(auth=auth)
    m = mail.get_messages_from_folder('Drafts')
    print('simple_by_folder {count}'.format(count=(len(m))))
    for message in m:
        print(message.Subject)


def inbox_parameters(auth):
    mail = Mail(auth=auth)
    filters = "From/EmailAddress/Address ne 'MicrosoftOffice365@email.office.com'" + \
              " and DateTimeReceived gt 2016-01-01T00:00:01Z"
    m = mail.inbox.get_messages(select=['DateTimeSent'],
                                filters=filters,
                                top=1)
    print('inbox_parameters {count}'.format(count=(len(m))))
    for message in m:
        print(message.Subject)


if __name__ == '__main__':
    authorization = (environ.get('OFFICE_USER'), environ.get('OFFICE_USER_PASSWORD'))
    simplest(authorization)
    simple_by_folder(authorization)
    inbox_parameters(authorization)

