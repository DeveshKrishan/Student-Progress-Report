<a name="readme-top"></a>
# Automated Student-Progress-Report

<!-- ABOUT THE PROJECT -->
## About

![image](https://github.com/DeveshKrishan/Student-Progress-Report/assets/91798447/dd832a78-50dd-4239-95c5-4b30e177f8d0)

The attached Google Sheet has the ability to download from Canvas (using GET API calls) to create a progress report of all students in a course with modules to which requirements are added. Remake of [Original](https://community.canvaslms.com/t5/Higher-Ed-Canvas-Users/Automated-progress-report-of-students-in-modules/ba-p/262284) in Google App Script.

<!-- GETTING STARTED -->
## Getting Started

To get a local copy up and running follow these simple steps.

![image](https://github.com/DeveshKrishan/Student-Progress-Report/assets/91798447/38ccbfcd-e626-4c9a-a912-e707be3fda36)

### Installation

_Note that this script relies upon the user is equivalent to the status of "Teacher" or above in the Canvas course._

1. Get a Canvas LMS API token by reading [Canvas LMS API Token Guide](https://community.canvaslms.com/t5/Admin-Guide/How-do-I-manage-API-access-tokens-as-an-admin/ta-p/89)
2. Copy the Google Sheet at [Example Sheet](https://docs.google.com/spreadsheets/d/1tzie8Ug5sHbygXiACGCmvYNAkXNHsp_OPe9x0ufpDP0/edit?usp=sharing)
3. Head to Google App Script to copy and paste the code from `code.js`
![image](https://github.com/DeveshKrishan/Student-Progress-Report/assets/91798447/c2ecef75-df42-46e8-82e8-42d972e0adda)


<!-- USAGE EXAMPLES -->
## Usage

Enter the course ID of the Canvas course as well as the API Token from Canvas LMS.

_Do note that you may have to input the Canvas URL of your institution for any URLs used to fetch data. Currently, this script is used for the University of California, Irvine Canvas Courses._

Click the Download from Canvas Button to begin running the script. 

![image](https://github.com/DeveshKrishan/Student-Progress-Report/assets/91798447/9dc4b9c2-dff4-4583-9017-ba9b196c7aad)

When you want to be ready to clear the data for any reason, feel free to use the Clear Button to erase any data downloaded. 

![image](https://github.com/DeveshKrishan/Student-Progress-Report/assets/91798447/b8322ff2-8c6d-45b0-80a2-0e26efd96fd7)


<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.txt` for more information.


<!-- ACKNOWLEDGMENTS -->
## Acknowledgments

! would like to acknowledge [stelpstra](https://community.canvaslms.com/t5/user/viewprofilepage/user-id/105030) for creating this script originally in Excel. This is a remake of his [product](https://community.canvaslms.com/t5/Higher-Ed-Canvas-Users/Automated-progress-report-of-students-in-modules/ba-p/262284) written in Google App Script. This project's design is heavily inspired by his original product. 


