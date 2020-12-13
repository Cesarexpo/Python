import os
import shutil

current_dir = os.getcwd()

for f in os.listdir(current_dir):
    filename, file_ext = os.path.splitext(f)

    try:
        if not file_ext:
            pass

        elif file_ext in ('.py'):
            shutil.move(
                os.path.join(current_dir, f'{filename}{file_ext}'),
                os.path.join(current_dir, 'Python files', f'{filename}{file_ext}'))
        elif file_ext in ('.jpg', '.png', '.gif'):
            shutil.move(
                os.path.join(current_dir, f'{filename}{file_ext}'),
                os.path.join(current_dir, 'Image files', f'{filename}{file_ext}'))
        elif file_ext in ('.xls', '.xlsx', '.xltx', '.xlsm'):
            shutil.move(
                os.path.join(current_dir, f'{filename}{file_ext}'),
                os.path.join(current_dir, 'Excel files', f'{filename}{file_ext}'))
        elif file_ext in ('.pdf'):
            shutil.move(
                os.path.join(current_dir, f'{filename}{file_ext}'),
                os.path.join(current_dir, 'PDF files', f'{filename}{file_ext}'))
        else:
            shutil.move(
                os.path.join(current_dir, f'{filename}{file_ext}'),
                os.path.join(current_dir, 'Other files', f'{filename}{file_ext}'))

    except (FileNotFoundError, PermissionError):
        pass
