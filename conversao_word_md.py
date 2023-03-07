import os.path
import win32com.client
import pypandoc

baseDir = 'stc2'

word = win32com.client.Dispatch("Word.application")

for dir_path, dirs, files in os.walk(baseDir):

	for file_name in files:

		file_path = os.path.join(dir_path, file_name)
		file_name, file_extension = os.path.splitext(file_path)

		if "~$" not in file_name:

			if file_extension == '.doc':

				docx_file = '{0}{1}'.format(file_path, 'x')

				if not os.path.isfile(docx_file):

					file_path = os.path.abspath(file_path)
					docx_file = os.path.abspath(docx_file)
					try:
						wordDoc = word.Documents.Open(file_path)
						wordDoc.SaveAs2(docx_file, FileFormat = 16)
						wordDoc.Close()
					except Exception as e:
						print('Failed to Convert: {0}'.format(file_path))
						print(e)

	for file_name in files:

		file_path = os.path.join(dir_path, file_name)
		file_name, file_extension = os.path.splitext(file_path)

		if "~$" not in file_name:

			if file_extension == '.docx':

				md_file = '{0}{1}'.format(file_name, '.md')

				if not os.path.isfile(md_file):

					print(file_path)
					print(md_file)

					file_path = os.path.abspath(file_path)
					pypandoc.convert_file(file_path, to='md', outputfile=md_file)