# Check Alumni/Students State
All MPI/MPU students can apply documents online (https://wapps2.ipm.edu.mo/rvdweb/pd_student_login.asp). This website requires the Identity Card and Date of Birth for entering the system. This is convenience for alumni because some may forget their student ID.

Actually, alumni can also find their student ID through the website when they forgot.

:star: This coding is for mass checking the current state of students.

## Required information
The following information should be sequentially write into the excel file (name.xlsx). For details, please reference the sample_name.xlsx.
- Identity Card number (format: 12345678, 1234567/8, 1234567(8))
- Month of Birth (format: two digits, e.g. 04)
- Day of Birth (format: two digits, e.g. 14)
![login](https://user-images.githubusercontent.com/34164281/162584816-56a5fea1-39ff-4e33-bd62-541893373c46.png)

## Getting the information from Website
After login in, the student ID, name and programmes are provided. With the understanding of the web design, some students/alumni information are included. We can check if that person is graduated or still leanring at MPI/MPU.

With the testing on different accounts, there are 3 situations about the faculty/school. Therefore, the faculty/school is not 100% correct, just for reference.
1. Alumni's accounts do not contain the faculty/school information
2. Undergrate students' account do not contain the faculty/school information
3. Currect graduate students contain the faculty information (if he/she also obtained another degree before, the school infomation of that degree is followed the new faculty information)

![sample_info](https://user-images.githubusercontent.com/34164281/162584434-7cd42584-11ad-4ad1-ad75-160f82a4c382.png)

## Saving checked information
Only saved the latest graduated information, no matter how many degrees that the students' had.

The checked infomation (graduated and non_graduated) is save into the excel file (with currect date identify).

## Sample files
In this repository, two sample files are provided. It is used for readers to understand what should be put in the name.xlsx for checking and what will be the output.

## Special Statement
- This coding is just for learning, please don't use it as other purpose.
- This design is available on 10-Apr 2022, if there is some changes in the future, this coding may not be workable.
