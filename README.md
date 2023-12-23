BIS 601 Assignment 6 – Access 3


BIS 601 Assignment #6 Access 3
Case: Management Development Institute (MDI)
Scheduling and Controlling Seminar Events
Objectives
In this case you will learn how to:
§ Create an Access database
§ Create a table
§ Design a table
§ Import data from Excel
§ Specify the primary key
§ Create a relationship between tables
§ Design a report
§ Use the Report Wizard
§ Calculate fields
§ Use aggregate functions
§ Use query summary features
Management Development Institute (MDI) was founded in 1984 by Harry D. Cougar to deliver quality
continuing education to managers and business leaders. Harry continues to serve as MDI’s President and
Chief Operating Officer, directing operations from the company’s modest headquarters located at the foot
of the Rocky Mountains in Golden, Colorado. MDI’s seminars are presented at a variety of locations
across North America, Europe, and Japan - wherever a management training need exists.
Daniel “Dan” MacFadden, MDI’s Vice President of Operations, has overall management responsibility
for developing seminars, scheduling seminars, recruiting and assigning instructors, arranging facilities,
and monitoring activities. For each seminar, Ingrid Olson, the Seminar Coordinator, arranges a facility
where the seminar is to be conducted. She frequently rents meeting space at popular and readily
accessible hotels in each city where a seminar is scheduled. Ingrid knows the importance of reserving
quality space, since an uncomfortable environment can have significant impact on the quality of a
seminar.
During their recent management retreat in Breckenridge, Harry, Dan, and Ingrid discussed the possibility
of establishing a database to provide them with better information for managing their operations. Dan
described to Harry and Ingrid how a Seminar Management System (SMS) could be used to address a
number of ad hoc questions concerning their seminar events, including those concerned with monitoring
and controlling the seat inventory. A seminar event (delivery) is when a seminar is held at a particular
location on a scheduled date. Ingrid described the embarrassment of MDI last week when she discovered
that a seminar was overbooked, because she did not have an accurate count of seats sold. Since MDI
highly values its customer relations, both Harry and Dan want to avoid this situation in the future. Harry
supported Ingrid’s concern for better inventory control information. It seemed that Dan’s proposal for a
SMS would meet their needs.
BIS 601 Assignment 6 – Access 3
Please don’t upload this file to any servers, or assignment collection websites. If you do you
violate CMU’s Academic Integrity Policies
2
Ingrid worked with Dan to design their database. They determined the data should be arranged into two
different tables. One that contains the data for a seminar; and another that has the data for an event, which
is the delivery of a particular seminar. With this design, they can enter the seminar data only once, and
then they can reference that data as they need it with each event. The result is their Entity Relation
Diagram (ERD) (Figure 1). Their ERD indicates there can be multiple events for each seminar. That is, a
one-to-many relationship exists between seminars and events.
Figure 1: Entity-Relation Diagram (ERD)
They continued with their design by placing the specific data that is unique for each seminar in one
database table (Figure 2), while the data specific to each event was placed in another database table
(Figure 3). The SeminarID in the event table acts as a foreign key linking the two tables.
FIELD NAME FIELD DESCRIPTION DATA TYPE FIELD SIZE
SeminarID Identification number uniquely
designating each seminar
Short Text 10
SeminarName Name of seminar Short Text 30
Program Program group (category) of seminar Short Text 5
Duration Duration of seminar in days Number Integer
Credits Number of Continuing Education Credits Number Single
Fee Seminar fee or price Currency (Automatic)
Figure 2: Seminar Table Structure
FIELD NAME FIELD DESCRIPTION DATA TYPE FIELD SIZE
EventID Identification number uniquely
specifying each offering of a seminar – a
seminar event
Number Integer
SeminarID Identification number uniquely
designating each seminar
Short Text 10
Location City where seminar event is to be held Short Text 30
Date Beginning date of seminar event Date/Time (Automatic)
Seats Total number of seats for seminar event
(event capacity)
Number Integer
Sold Number for seats sold for a seminar
event
Number Integer
ArrgComplete Whether or not facility arrangements are
completed for a seminar event
Short Text 3
Figure 3: Event Table Structure
You have been asked to perform the following tasks to build MDI’s Seminar Management System. When
finished, submit your database in Blackboard.
SEMINAR EVENT
BIS 601 Assignment 6 – Access 3
Please don’t upload this file to any servers, or assignment collection websites. If you do you
violate CMU’s Academic Integrity Policies
3
Task 1: Launch Access and create a new database that you name MDI.accdb. Create a new table, that
you name Seminar, which contains the fields from the Seminar table structure (Figure 2).
Specify the SeminarID as the primary key for this table. Save your database with your newly
defined table.
Task 2: Add the following records (rows) to the Seminar table:
SeminarID SeminarName Program Duration Credits Fee
C08104 COMMUNICATION SKILLS MGT 3 1.8 1415
C09231 EFFECTIVE EXECUTIVE SPEAK MGT 3 1.8 1195
F08314 FOREIGN EXCH STRATEGIES FIN 2 1.2 1125
F09101 FINANCIAL ANALYSIS FIN 4 2.5 1495
T10512 KNOWLEDGE MANAGEMENT MIS 3 1.8 1100
Save your Seminar table. You are now going to obtain the remaining records for the Seminar
table from an Excel file. This saves you the effort of typing this data. Open the MDI_Seminar
F22.xlxs file. Copy the seminar records to the Windows clipboard. Return to Access. Select
your Seminar table and then select the Datasheet View. Select the blank record below your last
entry in the table. Use the Paste Append command to include these records in your Seminar
table.
Task 3: Add the Event table to your MDI database. Use the External Data → Import Excel spreadsheet
command. The Events data is stored in the MDI_Event F22.xlxs file. After you have
imported the Event table from Excel, then use the Design View to specify the data type as
indicated in the Event Table Structure (Figure 3). Specify EventID as the primary key for this
table.
Task 4: Create the relationship between the Seminar table and the Event table as shown by the ERD
(Figure 1). Make sure to enforce all referential integrity constraints.
Task 5: Create an EventsList query using the Event and Seminar tables. This should include all the
database fields. However, there should be no duplicate field names in the query.
Task 6: Dan and Ingrid have several questions for which they want information. Do these queries to
help Ingrid obtain answers for those questions. She has prepared a Planning Analysis (Figure
4) to specify how to calculate fields needed in some of the queries.
Planning Analysis
Figure 4: Planning Analysis
Field Name Field Description Data Type Calculation
SeatsAvailable Available seats for an event Number Seats – Sold
RevSold Revenue of sold event seats Number Sold * Fee
RevSeats Revenue of Seats in each seminar event Number Seats * Fee
BIS 601 Assignment 6 – Access 3
Please don’t upload this file to any servers, or assignment collection websites. If you do you
violate CMU’s Academic Integrity Policies
4
The queries which Dan and Ingrid want you to answer are as follows:
a. Which seminar events are scheduled to be held in Denver? List the SeminarID,
SeminarName, and Date. Sort the results by Date in ascending order. Name the query
Task6A.
b. Which 2022 seminar events do not have the local arrangements completed? List the
EventID, Location, and Date. Sort the results by Date in ascending order. Name the
query Task6B.
c. What is the total revenue by seminar? List the SeminarID, SeminarName, and the
revenue of sold event seats (RevSold). Name the query Task6C.
d. Which seminar events are sold out? List the SeminarID, SeminarName, EventID, and
Date. Name the query Task6D.
e. Which MIS seminar events are available for a Fee of $1,100 or less? List the
SeminarName, Location, Date, and Fee. Sort the results by Date in ascending order.
Name the query Task6E.
f. What is the total possible revenue? List the revenue from all seats in all seminars
(RevSeats). Name the query Task6F.
g. What is the average seminar fee? Name the query Task6G.
h. Which seminars are the most/least popular? List the SeminarID, SeminarName, and the
average number of seats sold. Sort the results by the average number of seats sold in
descending order. Name the query Task6H.
i. What is the average duration of a seminar? Name the query Task6I.
j. How many 2022 seminar event seats remain to be sold? List the SeminarID,
SeminarName, EventID, Date, and the number of available seats (SeatsAvailable). Sort
the results by Date in ascending order. Make sure to only list those events which have
available seats. Name the query Task6J.
Task 7: Ingrid wants you to create a Seminar List report. This is a list of the current seminars that she
can send to potential attendees. The report should include the SeminarID, SeminarName,
Duration, Credits, and Fee fields from the Seminar table. Figure 5 shows a design sketch for
this report.
Management Development Institute
Seminar List
(current date)
Created by: Your Full Name and Class section
Name Number Length Credits Price
AFFIRMATIVE ACTION PLAN M09334 2 1.2 $995
CASH MANAGEMENT F09411 2 1.2 $1,345
COMMUNICATION SKILLS C08104 3 1.8 $1,415
.
.
.
Figure 5: Seminar List Report Design Sketch
BIS 601 Assignment 6 – Access 3
Please don’t upload this file to any servers, or assignment collection websites. If you do you
violate CMU’s Academic Integrity Policies
5
Task 8: Dan wants you to create a Seminar Schedule report. This is a list of the current seminar events
that are available from MDI. The report should include the SeminarName, Date, Location, and
EventID fields from the Seminar and Event table. Figure 6 shows a design sketch for this
report.
Management Development Institute
Seminar Schedule
(current date)
Created by: Your Full Name and Class section
Name Date Location EventID
AFFIRMATIVE ACTION PLAN 10/17/2022 SAN FRANCISCO 8111
12/24/2022 DENVER 8114
1/9/2023 LOS ANGELES 8112
3/15/2023 PHOENIX 8116
CASH MANAGEMENT 9/26/2022 CHICAGO 1278
12/1/2022 CHICAGO 1281
.
.
.
Figure 6: Seminar Schedule Report Design Sketch
Task 9: Dan wants you to create a short WebEx video, about 1 minute highlighting some of the tasks
you undertook. Upload a link to your video in Blackboard.
Access Features
Feature Help Lookup Key
Word
Help Lookup Topic
Aggregate functions Aggregate functions Display column totals in a datasheet →
Understand Sum and the other aggregate
functions
Calculate fields Calculate fields Calculated fields
Create database Create blank database Create a new database → Creating a
database without using a template
Create relationship Create relationship Create, edit or delete a relationship
Create table Create table Create a table
Import data from Excel Import Excel Import data from Excel to a new table in
Access
Primary key Primary key Create or remove a primary key
Report Wizard Report Wizard Create a simple report → Create a report
by suing the Report Wizard
