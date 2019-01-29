use trainingDb
----------------------------------------------------------------
create table BookDetails (
BookId int identity primary key,
Title varchar(20) not null,
Author varchar(50) null,
PublisherName varchar(50) null,
Category varchar(20) not null,
Price int,
ISBN varchar(20),
Borrowed bit
)

--select * from BookDetails
-----------------------------------------------------------------
create table MemberDetails (
StudentId int identity primary key,
FirstName varchar(20) not null,
MiddleName char null,
LastName varchar(20) not null,
Class varchar(10) not null,
Section varchar(10) not null,
Roll varchar(10) not null,
)

--select * from  MemberDetails
------------------------------------------------------------------
create table BookIssue(
StudentId int Foreign key references MemberDetails(StudentId),
StudentName varchar(50),
BookId int foreign key references BookDetails(BookId),
BookTitle varchar(100),
DateIssued date,
DateReturn date
)

--select * from bookissue
-------------------------------------------------------------------
create table BookReturn(
BookId int foreign key references BookDetails(BookId),
StudentId int Foreign key references MemberDetails(StudentId),
DateReturn date,
FineCollected int
)

--select * from BookReturn
------------------------------------------------------------------
--drop table setting
create table Setting(
id int identity,
daysNumber int,
fine int
)

--select * from setting
--insert into setting values(14,10)
