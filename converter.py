import pandas as pd


class converter():
    '''This class converts an excel file to a vcf file.'''

    def __init__(self, file_location):
        '''The file_location argument is the location of the excel file that will be converted.'''
        
        try :
            self.file = pd.read_excel(file_location)
        except:
            print("File not found")
            exit(1)

        self.file_as_list = self.file.values.tolist()

    def create_vcard_file(self, file_name='contacts.vcf', prefix='') -> None:
        '''This function creates a vcf file from the excel file.
        The file_name argument is the name of the vcf file that will be created.
        The prefix argument is the prefix that will be added to the names in the vcf file.
        
        Example:
        file_name = 'contacts.vcf'
        prefix = 'artitu'
        
        '''

        if file_name[-4:] != '.vcf':
            file_name += '.vcf'
            
        with open(file_name, 'w') as f:
            for row in self.file_as_list:
                row[0] = prefix + ' ' + row[0]
                row[1] = str(row[1])
                # TODO : fix edge cases of names and numbers
                    

                f.writelines(['BEGIN:VCARD\n', 'VERSION:4.0\n', 'N:', row[0], '\n', 'TEL;TYPE=CELL:', row[1], '\n', 'END:VCARD\n'])
            f.close()
        print("File created successfully")

        