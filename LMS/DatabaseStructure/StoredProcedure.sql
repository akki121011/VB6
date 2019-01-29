create proc usp_AddBook(
@Title varchar(20),
@Author varchar(50),
@PublisherName varchar(50),
@Category varchar(20),
@Price int,
@ISBN varchar(20),
@Borrowed bit
)
as 
begin
	begin try
	insert into BookDetails
	values (@Title,@Author,@PublisherName,@Category,@Price,@ISBN,@Borrowed)
	end try
	begin catch
	    SELECT ERROR_MESSAGE() AS ErrorMessage;
	end catch
end

--usp_AddBook 'RobinHood','New Robin','','Fiction',420,'12-134-1256',0
--------------------------------------------------------------------------------------
create proc usp_AddMember(
@FirstName varchar(20),
@MiddleName char ,
@LastName varchar(20),
@Class varchar(10),
@Section varchar(10),
@Roll varchar(10)
)
as 
begin
	begin try
	insert into MemberDetails
	values (@FirstName,@MiddleName,@LastName,@Class,@Section,@Roll)
	end try
	begin catch
	    SELECT ERROR_MESSAGE() AS ErrorMessage;
	end catch
end

--exec usp_AddMember 'Akash','','Gupta','X','A','123456'
-------------------------------------------------------------------------------------------
create proc usp_IssueBook(
@StudentId int,
@StudentName varchar(50),
@BookId int,
@BookTitle varchar(100),
@DateIssued date,
@DateReturn date
)
as 
begin
	begin try
	insert into bookissue
	values (@StudentId,@StudentName,@BookId,@BookTitle,@DateIssued,@DateReturn)
	end try
	begin catch
	    SELECT ERROR_MESSAGE() AS ErrorMessage;
	end catch
end

--select * from bookdetails
--select * from memberdetails
-----------------------------------------------------------------------------------------
create proc usp_GetIssueDate(
@studID int,
@bookId int)
as 
begin
	begin try
	select DateReturn from bookissue
	where StudentId=@studid and BookId=@bookid
	end try
	begin catch
	    SELECT ERROR_MESSAGE() AS ErrorMessage;
	end catch
end
--drop proc usp_GetIssueDate

--exec usp_GetIssueDate 1,1
-----------------------------------------------------------------------------------------
create proc usp_BookDisplay
as 
begin
	begin try
	select * from bookDetails
	end try
	begin catch
	    SELECT ERROR_MESSAGE() AS ErrorMessage;
	end catch
end
--drop proc usp_BookDisplay
-----------------------------------------------------------------------------------------
create proc usp_MemberDisplay
as 
begin
	begin try
	select * from memberdetails
	end try
	begin catch
	    SELECT ERROR_MESSAGE() AS ErrorMessage;
	end catch
end
------------------------------------------------------------
create proc usp_updateSetting(
@id as int,
@days as int,
@fine as int
)
as 
begin
	begin try
	update setting
	set daysNumber=@days,fine=@fine
	where id=@id
	end try
	begin catch
	    SELECT ERROR_MESSAGE() AS ErrorMessage;
	end catch
end

--drop proc usp_updateSetting
--exec usp_updateSetting 1,10,20
------------------------------------------------------------------------
create proc usp_DisplaySetting(
@id int)
as 
begin
	begin try
	select * from setting
	where id=@id
	end try
	begin catch
	    SELECT ERROR_MESSAGE() AS ErrorMessage;
	end catch
end

--exec usp_DisplaySetting 1