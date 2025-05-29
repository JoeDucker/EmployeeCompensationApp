USE master;
GO

IF EXISTS (SELECT name FROM sys.databases WHERE name = 'employee')
BEGIN
    ALTER DATABASE employee SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE employee;
END
GO

CREATE DATABASE employee;
GO

USE employee;
GO

IF EXISTS (SELECT name FROM sys.databases WHERE name = 'EmployeeDemoApp')
BEGIN
    ALTER DATABASE EmployeeDemoApp SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE EmployeeDemoApp;
END
GO


-- Table for Role
CREATE TABLE Role (
    Role_ID INT IDENTITY(1,1) PRIMARY KEY,
    Role_Name VARCHAR(100) UNIQUE NOT NULL
);

-- Table for Location
CREATE TABLE Location (
    Location_ID INT IDENTITY(1,1) PRIMARY KEY,
    Location_Name VARCHAR(100) UNIQUE NOT NULL
);

-- Table for Employee
CREATE TABLE Employee (
    Employee_ID VARCHAR(10) PRIMARY KEY,
    Name VARCHAR(100),
    Years_of_Experience FLOAT
);

-- Table for Employee Assignment
CREATE TABLE Employee_Assignment (
    Assignment_ID INT IDENTITY(1,1) PRIMARY KEY,
    Employee_ID VARCHAR(10),
    Role_ID INT,
    Location_ID INT,
    Active CHAR(1),
    Compensation INT,
    Last_Working_Day DATE,
    Reference_Date DATE,
    Period_From DATE,
    Period_To DATE,
    FOREIGN KEY (Employee_ID) REFERENCES Employee(Employee_ID),
    FOREIGN KEY (Role_ID) REFERENCES Role(Role_ID),
    FOREIGN KEY (Location_ID) REFERENCES Location(Location_ID)
);

-- Table for Employee Ratings
CREATE TABLE Employee_Rating (
    Employee_ID VARCHAR(10),
    Period_From DATE,
    Period_To DATE,
    Self_Rating FLOAT,
    Manager_Rating FLOAT,
    PRIMARY KEY (Employee_ID, Period_From, Period_To),
    FOREIGN KEY (Employee_ID) REFERENCES Employee(Employee_ID)
);

-- Table for Industry Compensation
CREATE TABLE Industry_Compensation (
    Location_ID INT,
    Role_ID INT,
    Average_Industry_Compensation INT,
    PRIMARY KEY (Location_ID, Role_ID),
    FOREIGN KEY (Location_ID) REFERENCES Location(Location_ID),
    FOREIGN KEY (Role_ID) REFERENCES Role(Role_ID)
);

INSERT INTO Role (Role_Name) VALUES ('Analyst');
INSERT INTO Role (Role_Name) VALUES ('Senior Analyst');
INSERT INTO Role (Role_Name) VALUES ('Associate');
INSERT INTO Role (Role_Name) VALUES ('Senior Associate');
INSERT INTO Role (Role_Name) VALUES ('Manager');


INSERT INTO Location (Location_Name) VALUES ('Jaipur');
INSERT INTO Location (Location_Name) VALUES ('Bangalore');
INSERT INTO Location (Location_Name) VALUES ('Pune');

INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C001', 'Aditi Khanna', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C002', 'Pooja Kaur', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C003', 'Kajal Mehta', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C004', 'Amaya Joshi', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C005', 'Bilal Kulkarni', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C006', 'Girish Reddy', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C007', 'Esha Dhillon', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C008', 'Kashi Nair', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C009', 'Himanshu Dhillon', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C010', 'Hetal Bhatia', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C011', 'Anusha Verma', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C012', 'Mukta Iyer', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C013', 'Hari Menon', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C014', 'Jai Menon', 6.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C015', 'Kalpana Sharma', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C016', 'Leena Verma', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C017', 'Arvind Sinha', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C018', 'Anshul Gupta', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C019', 'Devi Menon', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C020', 'Farhan Gupta', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C021', 'Rajeshwari Joshi', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C022', 'Damodar Kumar', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C023', 'Om Sinha', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C024', 'Kunal Kaur', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C025', 'Nishant Reddy', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C026', 'Gowri Iyer', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C027', 'Akshita Verma', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C028', 'Renu Gupta', 5.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C029', 'Mira Verma', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C030', 'Chhavi Mehta', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C031', 'Manasi Kulkarni', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C032', 'Kiran Banerjee', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C033', 'Farah Mehta', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C034', 'Kala Gupta', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C035', 'Aayush Verma', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C036', 'Akanksha Mehta', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C037', 'Lohit Reddy', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C038', 'Jayesh Chowdhury', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C039', 'Nirmala Verma', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C040', 'Alka Reddy', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C041', 'Praveen Joshi', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C042', 'Anisha Chowdhury', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C043', 'Prema Khanna', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C044', 'Pritha Sinha', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C045', 'Karthika Joshi', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C046', 'Kalyani Patel', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C047', 'Bindiya Pillai', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C048', 'Annapurna Mehta', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C049', 'Niraj Kumar', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C050', 'Nila Patel', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C051', 'Bhanu Kapoor', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C052', 'Deepali Joshi', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C053', 'Poornima Mehta', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C054', 'Manju Kaur', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C055', 'Rajesh Kapoor', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C056', 'Kamala Kumar', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C057', 'Reena Dhillon', 5.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C058', 'Gayathri Kumar', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C059', 'Archana Reddy', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C060', 'Devang Iyer', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C061', 'Rati Rao', 5.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C062', 'Bala Verma', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C063', 'Atul Bhatia', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C064', 'Astha Banerjee', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C065', 'Hima Kaur', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C066', 'Aarya Reddy', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C067', 'Ravi Bhatia', 6.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C068', 'Jyoti Kaur', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C069', 'Namrata Kulkarni', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C070', 'Abhilasha Kapoor', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C071', 'Anmol Dhillon', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C072', 'Ajit Dhillon', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C073', 'Nishita Kapoor', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C074', 'Nitesh Nair', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C075', 'Arun Iyer', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C076', 'Arjun Joshi', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C077', 'Nitya Menon', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C078', 'Indira Kumar', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C079', 'Hrithik Gupta', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C080', 'Nalini Sinha', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C081', 'Chandni Kaur', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C082', 'Avinash Dhillon', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C083', 'Inderjit Patel', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C084', 'Rani Khanna', 8.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C085', 'Akash Gupta', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C086', 'Anuj Sharma', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C087', 'Punit Bhatia', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C088', 'Mehul Kumar', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C089', 'Pratyush Kapoor', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C090', 'Harshita Chowdhury', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C091', 'Bhoomi Menon', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C092', 'Manish Chowdhury', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C093', 'Lata Sharma', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C094', 'Madhuri Nair', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C095', 'Poonam Dhillon', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C096', 'Anita Banerjee', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C097', 'Prem Menon', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C098', 'Ranjeet Kulkarni', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C099', 'Chinmay Gupta', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C100', 'Rahul Kumar', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C101', 'Abhijeet Nair', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C102', 'Chandan Bhatia', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C103', 'Hiral Mehta', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C104', 'Barkha Reddy', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C105', 'Radhika Sharma', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C106', 'Akhil Sharma', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C107', 'Ketan Pillai', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C108', 'Harini Khanna', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C109', 'Jagadish Nair', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C110', 'Mohini Joshi', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C111', 'Akshay Kumar', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C112', 'Isha Singh', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C113', 'Abhay Singh', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C114', 'Naina Khanna', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C115', 'Anil Pillai', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C116', 'Neeraj Dhillon', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C117', 'Lochana Singh', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C118', 'Meena Sharma', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C119', 'Nanda Pillai', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C120', 'Manisha Banerjee', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C121', 'Diya Rao', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C122', 'Kirti Rao', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C123', 'Malati Khanna', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C124', 'Kavya Menon', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C125', 'Agni Chowdhury', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C126', 'Damini Verma', 7.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C127', 'Amitabh Menon', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C128', 'Ishita Joshi', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C129', 'Drishti Bhatia', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C130', 'Jayant Kulkarni', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C131', 'Mayank Gupta', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C132', 'Rakesh Menon', 9.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C133', 'Raj Verma', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C134', 'Gita Kapoor', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C135', 'Pratibha Verma', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C136', 'Jitendra Bhatia', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C137', 'Balraj Singh', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C138', 'Dhaval Pillai', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C139', 'Leela Kumar', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C140', 'Ahana Banerjee', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C141', 'Mitra Singh', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C142', 'Manav Pillai', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C143', 'Ashok Chowdhury', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C144', 'Anjali Rao', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C145', 'Anand Sinha', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C146', 'Kamini Singh', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C147', 'Pallavi Pillai', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C148', 'Kushan Dhillon', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C149', 'Mahendra Iyer', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C150', 'Rajat Reddy', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C151', 'Jigar Rao', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C152', 'Nikita Sharma', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C153', 'Azad Kumar', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C154', 'Hema Banerjee', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C155', 'Dhanya Kulkarni', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C156', 'Ranjana Sinha', 9.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C157', 'Ayush Patel', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C158', 'Mohan Kapoor', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C159', 'Mukesh Nair', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C160', 'Chitrangada Patel', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C161', 'Durga Kaur', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C162', 'Megha Patel', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C163', 'Ayesha Sharma', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C164', 'Arnav Nair', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C165', 'Gopal Nair', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C166', 'Priyanka Chowdhury', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C167', 'Divya Banerjee', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C168', 'Purva Dhillon', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C169', 'Bipin Chowdhury', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C170', 'Dhruv Chowdhury', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C171', 'Prabhakar Gupta', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C172', 'Neha Mehta', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C173', 'Rajani Singh', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C174', 'Prajwal Kumar', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C175', 'Rashmi Chowdhury', 6.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C176', 'Gitanjali Joshi', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C177', 'Manjari Rao', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C178', 'Parvati Rao', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C179', 'Aanya Sharma', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C180', 'Haritha Kulkarni', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C181', 'Pratik Singh', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C182', 'Mansi Dhillon', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C183', 'Abishek Menon', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C184', 'Lalit Gupta', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C185', 'Nutan Khanna', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C186', 'Arushi Khanna', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C187', 'Ravindra Kaur', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C188', 'Amrita Khanna', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C189', 'Neela Bhatia', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C190', 'Rajni Iyer', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C191', 'Rajiv Nair', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C192', 'Ratan Banerjee', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C193', 'Ishani Kapoor', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C194', 'Keshav Sinha', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C195', 'Priti Kulkarni', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C196', 'Karan Reddy', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C197', 'Bhavin Joshi', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C198', 'Ashish Pillai', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C199', 'Ankita Kaur', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C200', 'Naresh Rao', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C201', 'Raghu Patel', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C202', 'Gayatri Verma', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C203', 'Indrani Verma', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C204', 'Rhea Sharma', 6.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C205', 'Anurag Kumar', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C206', 'Atharv Rao', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C207', 'Rachana Gupta', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C208', 'Alok Singh', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C209', 'Ishaan Reddy', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C210', 'Harsha Pillai', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C211', 'Deepa Reddy', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C212', 'Priya Pillai', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C213', 'Nitin Iyer', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C214', 'Lakshmi Mehta', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C215', 'Prutha Banerjee', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C216', 'Kamal Verma', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C217', 'Madhav Kapoor', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C218', 'Ankit Bhatia', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C219', 'Bijay Khanna', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C220', 'Kshitij Bhatia', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C221', 'Nandini Chowdhury', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C222', 'Ayaan Gupta', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C223', 'Pratima Reddy', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C224', 'Pankaj Chowdhury', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C225', 'Madhu Joshi', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C226', 'Mahima Menon', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C227', 'Deepika Nair', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C228', 'Amit Nair', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C229', 'Geetanjali Singh', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C230', 'Puneet Rao', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C231', 'Kedar Khanna', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C232', 'Bhavna Nair', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C233', 'Mukul Menon', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C234', 'Bikram Sinha', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C235', 'Amita Iyer', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C236', 'Amar Kapoor', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C237', 'Aishwarya Rao', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C238', 'Bhavya Iyer', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C239', 'Ajeet Kaur', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C240', 'Kavita Iyer', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C241', 'Preity Iyer', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C242', 'Avni Mehta', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C243', 'Abhiram Iyer', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C244', 'Narain Banerjee', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C245', 'Asha Kulkarni', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C246', 'Deepak Kapoor', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C247', 'Abhinav Joshi', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C248', 'Advaita Kulkarni', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C249', 'Ananya Kulkarni', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C250', 'Hemant Rao', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C251', 'Chetan Dhillon', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C252', 'Dhairya Khanna', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C253', 'Pavan Bhatia', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C254', 'Areen Kapoor', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C255', 'Kabir Dhillon', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C256', 'Parul Banerjee', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C257', 'Ila Sharma', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C258', 'Manjeet Bhatia', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C259', 'Aakash Patel', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C260', 'Rekha Mehta', 6.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C261', 'Prachi Sharma', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C262', 'Rashi Pillai', 5.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C263', 'Aditya Sinha', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C264', 'Avani Kaur', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C265', 'Padma Kulkarni', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C266', 'Aarav Gupta', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C267', 'Harish Sinha', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C268', 'Karthik Kapoor', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C269', 'Maruti Mehta', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C270', 'Fathima Sharma', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C271', 'Nitara Joshi', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C272', 'Nikhil Gupta', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C273', 'Darshan Singh', 8.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C274', 'Purab Kaur', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C275', 'Jaya Sinha', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C276', 'Jayanti Pillai', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C277', 'Gautam Patel', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C278', 'Manas Sinha', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C279', 'Lavanya Patel', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C280', 'Brinda Banerjee', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C281', 'Arundhati Menon', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C282', 'Advika Pillai', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C283', 'Akshara Patel', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C284', 'Anuradha Patel', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C285', 'Chaitanya Rao', 3.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C286', 'Nisha Singh', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C287', 'Neelam Kaur', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C288', 'Jhanvi Banerjee', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C289', 'Aravind Singh', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C290', 'Jaidev Khanna', 8.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C291', 'Chirag Sharma', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C292', 'Dhanush Sinha', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C293', 'Mitali Reddy', 1.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C294', 'Ketaki Kulkarni', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C295', 'Jahnavi Iyer', 5.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C296', 'Pradnya Patel', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C297', 'Khushi Chowdhury', 0.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C298', 'Preethi Nair', 2.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C299', 'Pushpa Mehta', 4.5);
INSERT INTO Employee (Employee_ID, Name, Years_of_Experience) VALUES ('C300', 'Ajay Bhatia', 0.5);

INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C001', 1, 1, 'Y', 653874, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C002', 2, 2, 'Y', 798703, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C003', 1, 2, 'Y', 663851, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C004', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C005', 3, 1, 'Y', 976125, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C006', 1, 3, 'Y', 670531, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C007', 1, 3, 'Y', 652076, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C008', 1, 2, 'Y', 655787, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C009', 2, 3, 'Y', 895396, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C010', 2, 3, 'Y', 820895, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C011', 1, 1, 'Y', 693712, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C012', 1, 2, 'Y', 680229, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C013', 1, 3, 'N', 570000, '2024-05-12', '2024-12-05', '2024-06-05', '2024-12-05');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C014', 4, 3, 'Y', 1704253, NULL, '2025-01-01', '2018-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C015', 1, 2, 'Y', 669692, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C016', 1, 2, 'Y', 677168, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C017', 2, 1, 'Y', 734868, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C018', 1, 1, 'Y', 667001, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C019', 4, 1, 'Y', 1714061, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C020', 1, 3, 'Y', 665610, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C021', 3, 2, 'Y', 1180331, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C022', 3, 1, 'Y', 910994, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C023', 2, 2, 'Y', 846412, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C024', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C025', 2, 2, 'N', 750000, '2024-05-15', '2024-05-15', '2021-11-14', '2024-05-15');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C026', 1, 3, 'Y', 696236, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C027', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C028', 4, 2, 'Y', 2291090, NULL, '2025-01-01', '2019-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C029', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C030', 3, 1, 'Y', 1032250, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C031', 1, 2, 'Y', 673339, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C032', 1, 2, 'N', 699114, '2024-06-25', '2024-06-25', '2022-12-25', '2024-06-25');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C033', 1, 3, 'N', 570000, '2024-06-25', '2024-06-25', '2023-12-25', '2024-06-25');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C034', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C035', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C036', 1, 1, 'Y', 695543, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C037', 1, 2, 'Y', 683759, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C038', 1, 2, 'N', 570000, '2024-05-14', '2024-05-14', '2023-11-13', '2024-05-14');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C039', 2, 2, 'Y', 839893, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C040', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C041', 2, 2, 'Y', 839782, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C042', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C043', 3, 2, 'Y', 1254653, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C044', 3, 2, 'Y', 1336002, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C045', 1, 2, 'N', 570000, '2024-05-17', '2024-05-17', '2023-11-16', '2024-05-17');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C046', 1, 2, 'N', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C047', 3, 1, 'Y', 1098051, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C048', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C049', 2, 2, 'N', 747748, '2024-06-01', '2024-01-06', '2021-07-07', '2024-01-06');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C050', 2, 2, 'Y', 748140, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C051', 2, 1, 'Y', 842482, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C052', 4, 1, 'Y', 1324502, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C053', 2, 2, 'Y', 738224, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C054', 1, 2, 'Y', 663969, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C055', 3, 2, 'Y', 1264754, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C056', 1, 2, 'Y', 669187, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C057', 4, 2, 'Y', 2452003, NULL, '2025-01-01', '2019-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C058', 1, 3, 'N', 570000, '2024-06-25', '2024-06-25', '2023-12-25', '2024-06-25');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C059', 1, 1, 'Y', 656199, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C060', 4, 1, 'Y', 1869885, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C061', 4, 2, 'Y', 1915774, NULL, '2025-01-01', '2019-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C062', 2, 1, 'Y', 787635, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C063', 2, 1, 'Y', 764485, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C064', 2, 1, 'Y', 810920, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C065', 2, 3, 'Y', 823326, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C066', 1, 1, 'Y', 668855, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C067', 4, 2, 'Y', 2007027, NULL, '2025-01-01', '2018-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C068', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C069', 1, 2, 'Y', 694768, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C070', 1, 1, 'Y', 658651, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C071', 1, 1, 'Y', 682058, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C072', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C073', 2, 2, 'N', 812000, '2024-06-25', '2024-06-25', '2021-12-25', '2024-06-25');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C074', 2, 2, 'Y', 806381, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C075', 2, 1, 'Y', 813200, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C076', 2, 1, 'Y', 723808, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C077', 2, 2, 'Y', 789598, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C078', 3, 3, 'Y', 907115, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C079', 2, 3, 'Y', 886891, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C080', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C081', 3, 1, 'N', 964851, '2024-06-25', '2024-06-25', '2021-12-25', '2024-06-25');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C082', 2, 1, 'Y', 765357, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C083', 3, 3, 'Y', 991402, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C084', 5, 2, 'Y', 3079338, NULL, '2025-01-01', '2016-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C085', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C086', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C087', 3, 2, 'N', 1093366, '2024-07-01', '2024-01-07', '2019-07-09', '2024-01-07');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C088', 1, 2, 'Y', 679372, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C089', 2, 2, 'Y', 882835, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C090', 2, 3, 'Y', 844734, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C091', 3, 1, 'Y', 974829, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C092', 1, 2, 'Y', 666771, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C093', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C094', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C095', 2, 2, 'N', 780000, '2024-07-01', '2024-01-07', '2021-07-08', '2024-01-07');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C096', 1, 1, 'Y', 674629, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C097', 3, 2, 'Y', 1161030, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C098', 4, 2, 'Y', 1744115, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C099', 3, 1, 'Y', 968562, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C100', 3, 2, 'Y', 1107472, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C101', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C102', 3, 1, 'Y', 1096816, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C103', 2, 3, 'N', 747378, '2024-06-01', '2024-01-06', '2021-07-07', '2024-01-06');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C104', 2, 1, 'Y', 784515, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C105', 3, 2, 'Y', 1038994, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C106', 1, 1, 'Y', 661872, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C107', 1, 2, 'N', 689119, '2024-05-12', '2024-12-05', '2023-06-06', '2024-12-05');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C108', 1, 3, 'Y', 679122, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C109', 4, 3, 'Y', 2119272, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C110', 1, 2, 'Y', 675832, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C111', 1, 1, 'Y', 677326, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C112', 3, 3, 'Y', 999847, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C113', 1, 1, 'Y', 667696, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C114', 1, 2, 'Y', 657556, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C115', 1, 1, 'Y', 666688, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C116', 2, 2, 'Y', 827191, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C117', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C118', 1, 2, 'Y', 650999, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C119', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C120', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C121', 1, 3, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C122', 1, 2, 'N', 570000, '2024-07-01', '2024-01-07', '2023-07-08', '2024-01-07');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C123', 1, 2, 'Y', 689746, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C124', 1, 2, 'N', 652764, '2024-06-25', '2024-06-25', '2022-12-25', '2024-06-25');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C125', 1, 1, 'Y', 668447, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C126', 5, 1, 'Y', 2910548, NULL, '2025-01-01', '2017-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C127', 1, 1, 'Y', 695121, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C128', 3, 3, 'Y', 990176, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C129', 1, 3, 'Y', 665297, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C130', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C131', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C132', 5, 2, 'Y', 3287201, NULL, '2025-01-01', '2015-07-05', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C133', 3, 2, 'Y', 1479593, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C134', 1, 3, 'N', 570000, '2024-06-25', '2024-06-25', '2023-12-25', '2024-06-25');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C135', 2, 2, 'Y', 866178, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C136', 1, 2, 'Y', 693467, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C137', 2, 1, 'N', 739278, '2024-05-12', '2024-12-05', '2022-06-06', '2024-12-05');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C138', 1, 3, 'Y', 688026, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C139', 1, 2, 'N', 570000, '2024-07-01', '2024-01-07', '2023-07-08', '2024-01-07');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C140', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C141', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C142', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C143', 2, 1, 'Y', 831898, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C144', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C145', 1, 1, 'Y', 675379, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C146', 1, 2, 'Y', 671736, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C147', 2, 2, 'Y', 823585, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C148', 1, 2, 'Y', 699616, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C149', 1, 2, 'Y', 693026, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C150', 3, 2, 'Y', 1580928, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C151', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C152', 2, 2, 'N', 800000, '2024-06-01', '2024-01-06', '2021-07-07', '2024-01-06');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C153', 2, 1, 'Y', 874157, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C154', 2, 3, 'Y', 837838, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C155', 1, 3, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C156', 5, 2, 'Y', 2586435, NULL, '2025-01-01', '2015-07-05', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C157', 2, 1, 'Y', 860479, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C158', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C159', 1, 2, 'N', 570000, '2024-05-12', '2024-12-05', '2024-06-05', '2024-12-05');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C160', 3, 1, 'Y', 1023034, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C161', 1, 3, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C162', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C163', 2, 1, 'Y', 797979, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C164', 2, 1, 'Y', 759937, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C165', 1, 3, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C166', 3, 2, 'Y', 1370487, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C167', 1, 3, 'Y', 694789, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C168', 3, 2, 'Y', 1022523, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C169', 3, 1, 'Y', 950310, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C170', 1, 3, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C171', 2, 2, 'Y', 878914, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C172', 2, 2, 'Y', 803203, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C173', 3, 2, 'Y', 1550000, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C174', 2, 2, 'Y', 798506, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C175', 4, 2, 'Y', 2415447, NULL, '2025-01-01', '2018-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C176', 1, 3, 'Y', 662617, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C177', 1, 2, 'Y', 685966, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C178', 2, 2, 'Y', 766540, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C179', 1, 1, 'Y', 684097, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C180', 2, 3, 'Y', 874620, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C181', 2, 2, 'Y', 894749, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C182', 1, 2, 'N', 570000, '2024-07-01', '2024-01-07', '2023-07-08', '2024-01-07');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C183', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C184', 1, 2, 'Y', 688925, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C185', 2, 2, 'N', 757860, '2024-05-12', '2024-12-05', '2022-06-06', '2024-12-05');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C186', 2, 1, 'Y', 731274, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C187', 4, 2, 'Y', 2084629, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C188', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C189', 2, 2, 'Y', 781187, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C190', 3, 2, 'N', 1100000, '2024-07-01', '2024-01-07', '2019-07-09', '2024-01-07');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C191', 3, 2, 'Y', 1364647, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C192', 4, 2, 'Y', 2078123, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C193', 3, 3, 'Y', 1011693, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C194', 1, 2, 'N', 660441, '2024-05-17', '2024-05-17', '2022-11-16', '2024-05-17');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C195', 3, 2, 'Y', 1171667, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C196', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C197', 2, 1, 'Y', 800808, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C198', 2, 1, 'Y', 808913, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C199', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C200', 1, 2, 'Y', 669647, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C201', 3, 2, 'Y', 1279950, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C202', 1, 3, 'Y', 667714, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C203', 3, 3, 'N', 1100000, '2024-05-12', '2024-12-05', '2020-06-06', '2024-12-05');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C204', 4, 2, 'Y', 1884348, NULL, '2025-01-01', '2018-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C205', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C206', 2, 1, 'N', 850060, '2024-07-01', '2024-01-07', '2021-07-08', '2024-01-07');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C207', 3, 2, 'N', 1150000, '2024-07-01', '2024-01-07', '2020-07-08', '2024-01-07');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C208', 1, 1, 'Y', 679088, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C209', 3, 3, 'Y', 901372, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C210', 2, 3, 'N', 743907, '2024-05-12', '2024-12-05', '2022-06-06', '2024-12-05');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C211', 4, 1, 'Y', 1324502, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C212', 3, 2, 'Y', 1423630, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C213', 2, 2, 'Y', 822884, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C214', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C215', 3, 2, 'Y', 1130975, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C216', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C217', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C218', 1, 1, 'Y', 682722, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C219', 3, 1, 'Y', 958924, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C220', 1, 2, 'Y', 698795, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C221', 1, 2, 'Y', 673115, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C222', 2, 1, 'Y', 874505, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C223', 2, 2, 'N', 757280, '2024-05-12', '2024-12-05', '2022-06-06', '2024-12-05');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C224', 2, 2, 'Y', 885800, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C225', 1, 2, 'Y', 681109, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C226', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C227', 4, 1, 'Y', 1542655, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C228', 1, 1, 'Y', 669794, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C229', 1, 3, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C230', 3, 2, 'Y', 1502831, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C231', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C232', 2, 1, 'Y', 881213, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C233', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C234', 3, 1, 'Y', 1020947, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C235', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C236', 1, 1, 'Y', 687204, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C237', 1, 1, 'Y', 696134, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C238', 2, 1, 'Y', 807821, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C239', 1, 1, 'Y', 652519, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C240', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C241', 2, 2, 'Y', 749271, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C242', 2, 1, 'Y', 813458, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C243', 1, 1, 'Y', 695357, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C244', 1, 2, 'N', 570000, '2024-07-01', '2024-01-07', '2023-07-08', '2024-01-07');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C245', 2, 1, 'Y', 808047, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C246', 4, 1, 'Y', 1714061, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C247', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C248', 1, 1, 'Y', 694539, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C249', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C250', 2, 3, 'Y', 732142, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C251', 3, 1, 'Y', 917419, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C252', 4, 1, 'Y', 1402414, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C253', 2, 2, 'N', 865555, '2024-05-17', '2024-05-17', '2021-11-16', '2024-05-17');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C254', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C255', 1, 2, 'N', 570000, '2024-05-17', '2024-05-17', '2023-11-16', '2024-05-17');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C256', 2, 2, 'Y', 753470, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C257', 2, 3, 'Y', 805979, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C258', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C259', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C260', 4, 2, 'Y', 2121197, NULL, '2025-01-01', '2018-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C261', 2, 2, 'Y', 816659, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C262', 4, 2, 'Y', 2295146, NULL, '2025-01-01', '2019-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C263', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C264', 2, 1, 'Y', 813559, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C265', 2, 2, 'Y', 879923, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C266', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C267', 1, 3, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C268', 1, 2, 'Y', 660243, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C269', 1, 2, 'Y', 693416, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C270', 1, 3, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C271', 2, 2, 'Y', 754564, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C272', 2, 2, 'Y', 856722, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C273', 5, 1, 'Y', 3039209, NULL, '2025-01-01', '2016-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C274', 3, 2, 'Y', 1137127, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C275', 1, 2, 'Y', 657432, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C276', 1, 2, 'Y', 672615, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C277', 1, 3, 'Y', 686097, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C278', 1, 2, 'N', 570000, '2024-05-17', '2024-05-17', '2023-11-16', '2024-05-17');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C279', 1, 2, 'Y', 659554, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C280', 3, 1, 'Y', 1019658, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C281', 2, 1, 'Y', 841886, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C282', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C283', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C284', 1, 1, 'Y', 654800, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C285', 3, 1, 'Y', 1031403, NULL, '2025-01-01', '2021-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C286', 2, 2, 'Y', 842589, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C287', 2, 2, 'Y', 771600, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C288', 1, 2, 'Y', 689961, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C289', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C290', 5, 3, 'Y', 2945176, NULL, '2025-01-01', '2016-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C291', 3, 1, 'Y', 1007064, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C292', 4, 1, 'Y', 1402414, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C293', 1, 2, 'Y', 664724, NULL, '2025-01-01', '2023-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C294', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C295', 4, 3, 'Y', 1729701, NULL, '2025-01-01', '2019-07-04', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C296', 2, 2, 'Y', 882632, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C297', 1, 2, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C298', 2, 2, 'Y', 828900, NULL, '2025-01-01', '2022-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C299', 3, 2, 'Y', 1030381, NULL, '2025-01-01', '2020-07-03', '2025-01-01');
INSERT INTO Employee_Assignment (Employee_ID, Role_ID, Location_ID, Active, Compensation, Last_Working_Day, Reference_Date, Period_From, Period_To) VALUES ('C300', 1, 1, 'Y', 570000, NULL, '2025-01-01', '2024-07-02', '2025-01-01');


INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C001', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C002', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C003', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C004', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C005', '2021-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C006', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C007', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C008', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C009', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C010', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C011', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C012', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C013', '2024-06-05', '2024-12-05', 5.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C014', '2018-07-04', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C015', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C016', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C017', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C018', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C019', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C020', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C021', '2021-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C022', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C023', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C024', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C025', '2021-11-14', '2024-05-15', 4.5, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C026', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C027', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C028', '2019-07-04', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C029', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C030', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C031', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C032', '2022-12-25', '2024-06-25', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C033', '2023-12-25', '2024-06-25', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C034', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C035', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C036', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C037', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C038', '2023-11-13', '2024-05-14', 5.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C039', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C040', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C041', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C042', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C043', '2021-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C044', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C045', '2023-11-16', '2024-05-17', 5.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C046', '2024-07-02', '2025-01-01', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C047', '2020-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C048', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C049', '2021-07-07', '2024-01-06', 5.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C050', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C051', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C052', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C053', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C054', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C055', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C056', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C057', '2019-07-04', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C058', '2023-12-25', '2024-06-25', 5.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C059', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C060', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C061', '2019-07-04', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C062', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C063', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C064', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C065', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C066', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C067', '2018-07-04', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C068', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C069', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C070', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C071', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C072', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C073', '2021-12-25', '2024-06-25', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C074', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C075', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C076', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C077', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C078', '2021-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C079', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C080', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C081', '2021-12-25', '2024-06-25', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C082', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C083', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C084', '2016-07-04', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C085', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C086', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C087', '2019-07-09', '2024-01-07', 5.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C088', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C089', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C090', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C091', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C092', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C093', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C094', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C095', '2021-07-08', '2024-01-07', 5.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C096', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C097', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C098', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C099', '2021-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C100', '2020-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C101', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C102', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C103', '2021-07-07', '2024-01-06', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C104', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C105', '2020-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C106', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C107', '2023-06-06', '2024-12-05', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C108', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C109', '2020-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C110', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C111', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C112', '2021-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C113', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C114', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C115', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C116', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C117', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C118', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C119', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C120', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C121', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C122', '2023-07-08', '2024-01-07', 5.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C123', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C124', '2022-12-25', '2024-06-25', 5.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C125', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C126', '2017-07-04', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C127', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C128', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C129', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C130', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C131', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C132', '2015-07-05', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C133', '2021-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C134', '2023-12-25', '2024-06-25', 5.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C135', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C136', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C137', '2022-06-06', '2024-12-05', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C138', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C139', '2023-07-08', '2024-01-07', 5.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C140', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C141', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C142', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C143', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C144', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C145', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C146', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C147', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C148', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C149', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C150', '2021-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C151', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C152', '2021-07-07', '2024-01-06', 5.0, 3.5);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C153', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C154', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C155', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C156', '2015-07-05', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C157', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C158', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C159', '2024-06-05', '2024-12-05', 5.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C160', '2021-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C161', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C162', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C163', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C164', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C165', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C166', '2021-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C167', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C168', '2021-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C169', '2021-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C170', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C171', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C172', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C173', '2020-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C174', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C175', '2018-07-04', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C176', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C177', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C178', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C179', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C180', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C181', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C182', '2023-07-08', '2024-01-07', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C183', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C184', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C185', '2022-06-06', '2024-12-05', 5.0, 3.5);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C186', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C187', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C188', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C189', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C190', '2019-07-09', '2024-01-07', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C191', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C192', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C193', '2021-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C194', '2022-11-16', '2024-05-17', 5.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C195', '2021-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C196', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C197', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C198', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C199', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C200', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C201', '2021-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C202', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C203', '2020-06-06', '2024-12-05', 5.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C204', '2018-07-04', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C205', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C206', '2021-07-08', '2024-01-07', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C207', '2020-07-08', '2024-01-07', 5.0, 2.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C208', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C209', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C210', '2022-06-06', '2024-12-05', 5.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C211', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C212', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C213', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C214', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C215', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C216', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C217', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C218', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C219', '2021-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C220', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C221', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C222', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C223', '2022-06-06', '2024-12-05', 5.0, 3.5);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C224', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C225', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C226', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C227', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C228', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C229', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C230', '2021-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C231', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C232', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C233', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C234', '2020-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C235', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C236', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C237', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C238', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C239', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C240', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C241', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C242', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C243', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C244', '2023-07-08', '2024-01-07', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C245', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C246', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C247', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C248', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C249', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C250', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C251', '2021-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C252', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C253', '2021-11-16', '2024-05-17', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C254', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C255', '2023-11-16', '2024-05-17', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C256', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C257', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C258', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C259', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C260', '2018-07-04', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C261', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C262', '2019-07-04', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C263', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C264', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C265', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C266', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C267', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C268', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C269', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C270', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C271', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C272', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C273', '2016-07-04', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C274', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C275', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C276', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C277', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C278', '2023-11-16', '2024-05-17', 4.0, 3.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C279', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C280', '2020-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C281', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C282', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C283', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C284', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C285', '2021-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C286', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C287', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C288', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C289', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C290', '2016-07-04', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C291', '2020-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C292', '2023-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C293', '2023-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C294', '2024-07-02', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C295', '2019-07-04', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C296', '2022-07-03', '2025-01-01', 5.0, 5.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C297', '2024-07-02', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C298', '2022-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C299', '2020-07-03', '2025-01-01', 4.0, 4.0);
INSERT INTO Employee_Rating (Employee_ID, Period_From, Period_To, Self_Rating, Manager_Rating) VALUES ('C300', '2024-07-02', '2025-01-01', 5.0, 5.0);

INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (1, 5, 2337356);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (1, 4, 1558238);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (1, 3, 1038825);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (1, 2, 769500);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (1, 1, 570000);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (3, 5, 2583394);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (3, 4, 1722263);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (3, 3, 1148175);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (3, 2, 850500);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (3, 1, 630000);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (2, 5, 2706413);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (2, 4, 1804275);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (2, 3, 1202850);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (2, 2, 891000);
INSERT INTO Industry_Compensation (Location_ID, Role_ID, Average_Industry_Compensation) VALUES (2, 1, 660000);

-- User Story 1: Filter and Display Active Employees by Role -- 
-- Filter employees by Role --
-- All Roles --

SELECT 
    e.Name, r.Role_Name, l.Location_Name, ea.Compensation, ea.Active
FROM Employee_Assignment ea
JOIN Employee e ON ea.Employee_ID = e.Employee_ID
JOIN Role r ON ea.Role_ID = r.Role_ID
JOIN Location l ON ea.Location_ID = l.Location_ID;

-- All active employees, across all roles and all locations --

SELECT 
    e.Name AS Employee_Name,
    r.Role_Name,
    l.Location_Name,
    ea.Compensation
FROM Employee_Assignment ea
JOIN Employee e ON ea.Employee_ID = e.Employee_ID
JOIN Role r ON ea.Role_ID = r.Role_ID
JOIN Location l ON ea.Location_ID = l.Location_ID
WHERE ea.Active = 'Y';

-- Average compensation for all location --

SELECT 
    l.Location_Name,
    AVG(ea.Compensation) AS Avg_Compensation
FROM Employee_Assignment ea
JOIN Location l ON ea.Location_ID = l.Location_ID
WHERE ea.Active = 'Y'
GROUP BY l.Location_Name;

-- Toggle to include/exclude Inactive employees --

SELECT 
    e.Name AS Employee_Name,
    r.Role_Name,
    l.Location_Name,
    ea.Compensation
FROM Employee_Assignment ea
JOIN Employee e ON ea.Employee_ID = e.Employee_ID
JOIN Role r ON ea.Role_ID = r.Role_ID
JOIN Location l ON ea.Location_ID = l.Location_ID
WHERE ea.Active = 'Y';

-- Attrition Rate for Bangalore, Pune, and Jaipur --

SELECT 
    l.Location_Name,
    COUNT(CASE WHEN ea.Active = 'N' THEN 1 END) AS Inactive_Count,
    COUNT(*) AS Total_Employees,
    ROUND(100.0 * COUNT(CASE WHEN ea.Active = 'N' THEN 1 END) / COUNT(*), 2) AS Attrition_Rate_Percentage
FROM Employee_Assignment ea
JOIN Location l ON ea.Location_ID = l.Location_ID
WHERE l.Location_Name IN ('Bangalore', 'Pune', 'Jaipur')
GROUP BY l.Location_Name;

-- User Story 2: Group Employees by Years of Experience --
-- Count Employees by Experience Ranges (Using Actual Values)--

SELECT
    CASE
        WHEN e.Years_of_Experience < 1 THEN '0–1 years'
        WHEN e.Years_of_Experience < 2 THEN '1–2 years'
        WHEN e.Years_of_Experience < 3 THEN '2–3 years'
        WHEN e.Years_of_Experience < 4 THEN '3–4 years'
        WHEN e.Years_of_Experience < 5 THEN '4–5 years'
        WHEN e.Years_of_Experience < 6 THEN '5–6 years'
        WHEN e.Years_of_Experience < 7 THEN '6–7 years'
        WHEN e.Years_of_Experience < 8 THEN '7–8 years'
        WHEN e.Years_of_Experience < 9 THEN '8–9 years'
        WHEN e.Years_of_Experience < 10 THEN '9–10 years'
        ELSE '10+ years'
    END AS Experience_Range,
    l.Location_Name,
    r.Role_Name,
    COUNT(*) AS Employee_Count
FROM Employee e
JOIN Employee_Assignment ea ON e.Employee_ID = ea.Employee_ID
JOIN Role r ON ea.Role_ID = r.Role_ID
JOIN Location l ON ea.Location_ID = l.Location_ID
WHERE ea.Active = 'Y'
GROUP BY
    CASE
        WHEN e.Years_of_Experience < 1 THEN '0–1 years'
        WHEN e.Years_of_Experience < 2 THEN '1–2 years'
        WHEN e.Years_of_Experience < 3 THEN '2–3 years'
        WHEN e.Years_of_Experience < 4 THEN '3–4 years'
        WHEN e.Years_of_Experience < 5 THEN '4–5 years'
        WHEN e.Years_of_Experience < 6 THEN '5–6 years'
        WHEN e.Years_of_Experience < 7 THEN '6–7 years'
        WHEN e.Years_of_Experience < 8 THEN '7–8 years'
        WHEN e.Years_of_Experience < 9 THEN '8–9 years'
        WHEN e.Years_of_Experience < 10 THEN '9–10 years'
        ELSE '10+ years'
    END,
    l.Location_Name,
    r.Role_Name
ORDER BY
    CASE
        WHEN e.Years_of_Experience < 1 THEN '0–1 years'
        WHEN e.Years_of_Experience < 2 THEN '1–2 years'
        WHEN e.Years_of_Experience < 3 THEN '2–3 years'
        WHEN e.Years_of_Experience < 4 THEN '3–4 years'
        WHEN e.Years_of_Experience < 5 THEN '4–5 years'
        WHEN e.Years_of_Experience < 6 THEN '5–6 years'
        WHEN e.Years_of_Experience < 7 THEN '6–7 years'
        WHEN e.Years_of_Experience < 8 THEN '7–8 years'
        WHEN e.Years_of_Experience < 9 THEN '8–9 years'
        WHEN e.Years_of_Experience < 10 THEN '9–10 years'
        ELSE '10+ years'
    END,
    l.Location_Name,
    r.Role_Name;


-- User Story 3: Simulate Compensation Increments --
-- Global % Increment Simulation (for all active employees) --

SELECT 
    e.Name AS Employee_Name,
    r.Role_Name,
    l.Location_Name,
    ea.Compensation AS Current_Compensation,
    ROUND(ea.Compensation * (1 + 0.10), 2) AS Updated_Compensation_10pct
FROM Employee_Assignment ea
JOIN Employee e ON ea.Employee_ID = e.Employee_ID
JOIN Role r ON ea.Role_ID = r.Role_ID
JOIN Location l ON ea.Location_ID = l.Location_ID
WHERE ea.Active = 'Y';

-- Bonus: Custom % Increment per Location --

SELECT 
    e.Name AS Employee_Name,
    l.Location_Name,
    r.Role_Name,
    ea.Compensation AS Current_Compensation,
    ROUND(
        ea.Compensation * 
        (1 + 
            CASE 
                WHEN l.Location_Name = 'Bangalore' THEN 0.10
                WHEN l.Location_Name = 'Pune' THEN 0.12
                WHEN l.Location_Name = 'Jaipur' THEN 0.08
                ELSE 0.00
            END
        ), 2
    ) AS Updated_Compensation
FROM Employee_Assignment ea
JOIN Employee e ON ea.Employee_ID = e.Employee_ID
JOIN Role r ON ea.Role_ID = r.Role_ID
JOIN Location l ON ea.Location_ID = l.Location_ID
WHERE ea.Active = 'Y';

-- User Story 4: Download Filtered Employee Data --
-- Export filtered employee data to a CSV file -- 

SELECT 
    e.Name AS Employee_Name,
    r.Role_Name,
    l.Location_Name,
    ea.Compensation AS Current_Compensation,
    ROUND(ea.Compensation * 1.10, 2) AS Updated_Compensation
FROM Employee_Assignment ea
JOIN Employee e ON ea.Employee_ID = e.Employee_ID
JOIN Role r ON ea.Role_ID = r.Role_ID
JOIN Location l ON ea.Location_ID = l.Location_ID
WHERE ea.Active = 'Y';
