import mysql.connector

#Construtor for establishing the connection.
class dataset:
    def __init__(self,host='localhost',username='root',password='sharvi@123',database='Student'):
       self.con = mysql.connector.connect(host=host,username=username,
                                     password=password,database=database)
       self.cursor = self.con.cursor()
    #function for adding data
       
    def add_data(self,roll,name,email,age,gender,contact,address):
        self.roll = roll
        self.name = name
        self.email = email
        self.age = age
        self.gender = gender
        self.contact = contact
        self.address = address
        query = 'insert into Student.Students(Roll_No,Name,Email,Age,Gender,Contact,Address) values(%s,%s,%s,%s,%s,%s,%s)'
        record = (roll,name,email,age,gender,contact,address)
        self.cursor.execute(query,record)
        self.con.commit()

    #function for retriving data
    def retrive_data(self):
        query = 'select * from Student.Students'
        self.cursor.execute(query)
        result = self.cursor.fetchall()
        return result
    # function for search data w.r.t roll 
    def search_roll_data(self,temp):
        query = """ select * from Student.Students where Roll_No = %s"""
        tup = (temp,)
        self.cursor.execute(query,tup)
        result = self.cursor.fetchall()
        return result
    # function for search data w.r.t Name
    def search_name_data(self,temp):
        query = "select * from Student.Students where Name LIKE '%" + str(temp) + "%'"
        # tup = (temp,)
        self.cursor.execute(query)
        result = self.cursor.fetchall()
        return result
    # function for search data w.r.t contact no
    def search_contact_data(self,temp):
        query = """ select * from Student.Students where Contact = %s"""
        tup = (temp,)
        self.cursor.execute(query,tup)
        result = self.cursor.fetchall()
        return result
    # function for search data w.r.t gender
    def search_gender_data(self,temp):
        query = """ select * from Student.Students where Gender = %s"""
        tup = (temp,)
        self.cursor.execute(query,tup)
        result = self.cursor.fetchall()
        return result
    # function for updating data
    def update_data(self, roll, new_name, new_email, new_age, new_gender, new_contact, new_address):
        try:
            # Validate data types
            int(roll)
            int(new_age)

            # Build the query with placeholders for security
            query = """
                UPDATE Student.Students
                SET Name = %s, Email = %s, Age = %s, Gender = %s, Contact = %s, Address = %s
                WHERE Roll_No = %s
            """
            record = (
                new_name,
                new_email,
                new_age,
                new_gender,
                new_contact,
                new_address,
                roll,
            )

            self.cursor.execute(query, record)
            self.con.commit()
            print("Record updated successfully!")

        except (ValueError, mysql.connector.Error) as e:
            print("Error updating data:", e)

    #function for deleting data.
    def delete_data(self, roll):
        try:
            # Validate roll number data type
            int(roll)

            # Build the delete query with a placeholder for security
            query = "DELETE FROM Student.Students WHERE Roll_No = %s"
            record = (roll,)

            self.cursor.execute(query, record)
            self.con.commit()
            print("Record deleted successfully!")

        except ValueError:
            print("Invalid roll number. Please enter an integer.")
        except mysql.connector.Error as e:
            print("Error deleting data:", e)

    def close_connection(self):
        """Closes the connection to the database."""
        if self.con:
            self.con.close()