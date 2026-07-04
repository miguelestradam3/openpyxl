class DataManager():
    #module as attribute to facilitate imports
    pd = __import__('pandas') #pandas module as attribute
    sys = __import__('sys') #sys module as attribute
    sqlite3 = __import__('sqlite3') #sqlite3 module as attribute

    #normal attributes
    filename_ = '' #excel output filename
    database_ = '' #database file location

    #------------------------
    # Methods
    #------------------------

    def __init__(self):
        """
        Function that initializes class
        Objective: flag the start of Data Handling
        Params: No arguments/parameters
        """
        print('Starting data management services...')

    @property
    def filename(self):
        """
        Getter method
        ---
        Objective: retrieve excel filename
        Params: No arguments/parameters
        """
        return self.filename_

    @filename.setter
    def filename(self, filename:str):
        """
        Getter method
        ---
        Objective: Save data into filename provided
        Params: filename (string [Required])
        """
        print('setting new data source...')
        self.filename_ = filename
        print('reading source...')
        try:
            self.xls = self.pd.ExcelFile(self.filename)
        except Exception as error:
            print('Unable to read Excel file...')
            print("\nERROR: {}\n".format(error))
            self.sys.exit(-1)

    @property
    def database(self):
        """
        Getter method
        ---
        Objective: returns database path location.
        Params: No arguments/parameters
        """
        return self.database_

    @database.setter
    def database(self, db_path:str):
        """
        Setter method
        ---
        Objective: Called to define the connection to SqLite connection
        Params:
            param -> db_path: Location of database into which we are going to write/get our data.
            param -> db_path: string       
        """
        try:
            self.database_ = db_path
            print('connecting to database...')
            self.connection = self.sqlite3.connect(self.database)
            self.cursor = self.connection.cursor()
            print('connection to database started...')
        except Exception as error:
            print("\nERROR: {}\n".format(error))

    def run_query(self, from_file:bool=False, script:str="SELECT * FROM Invoices")->None:
        """
        Setter method
        ---
        Objective: Called to define the connection to SqLite connection
        Params:
            param -> db_path -> description: 
            param -> db_path -> type: string
            param -> db_path -> default: 'SHOW TABLES;'

            param -> db_path -> description: 
            param -> db_path -> type: bool
            param -> db_path -> default: False
        """
        try:
            if from_file:
                with open(script, 'r') as sql_file:
                    script = sql_file.read()
            self.cursor.execute(script)
            print('\n',self.cursor.fetchall(),'\n')
        except Exception as error:
            print("\nERROR: {}\n".format(error))
        finally:
            if self.save_changes():
                print('commited action...')

    def save_changes(self)->bool:
        """
        Class method
        ---
        Output: boolean value returned
        Params: No arguments/parameters
        Objective: Changes into SqLite database have to be commited in order to be saved.  
        """
        try:
            print('Automatically saving changes...')
            self.connection.commit()
            return True
        except Exception as error:
            print("\nERROR: {}\n".format(error))
        return False


    def excel_to_db(self)->None:
        """
        Class method
        ---
        Params: No parameters/arguments
        Objective: Write database to excel sheet with tables as sheets.
        """
        try:
            assert self.filename != ''
            assert self.connection != ''
            print('transforming excel into SqLite Database')
            for sheet_name in self.xls.sheet_names:
                try:
                    df = self.pd.read_excel(self.xls, sheet_name=sheet_name)
                    df.to_sql(sheet_name, self.connection, if_exists="replace")
                    print('reading sheet {}...'.format(sheet_name))
                except:
                    print('ERROR: Unable to read sheet {} from excel file...'.format(sheet_name))
        except Exception as error:
            print("\nERROR: {}\n".format(error))

    def all_db_to_excel(self)->None:
        """
        Class method
        ---
        Params: No parameters/arguments
        Objective: Write excel file from a Sqlite database
        """
        # Get all table names
        tables = self.pd.read_sql_query(
            "SELECT name FROM sqlite_master WHERE type='table';",
            self.connection
        )
        # Create an Excel workbook
        with self.pd.ExcelWriter(self.filename_, engine="openpyxl") as writer:

            for table in tables["name"]:
                df = self.pd.read_sql_query(f"SELECT * FROM {table}", self.connection)

                # Excel sheet names have a maximum length of 31 characters
                df.to_excel(writer, sheet_name=table[:31], index=False)
        print("Database exported successfully!")

    def __del__(self):
        """
        Deletion method, this is done automatically once all the tasks have been executed
        Objective: Marks the end of processing.
        Params: No arguments/parameters
        """
        self.save_changes()

        self.cursor.close()
        print('finished reading and transforming data...')