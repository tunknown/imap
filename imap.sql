-- TUnknown 2020 license cc0/public domain
-- https://en.wikipedia.org/wiki/Creative_Commons_license#Zero_/_public_domain
-- ��� �������� ����� ������ CLR ����� �������, ����� ���� ���������� ������

-- http://www.t-sql.ru/post/IMAPClr.aspx
-- Admin:
--	if process is hung then to stop it can be close 143 port connection via tcpvew.exe on sql server
--	delete invalid mail from INBOX


EXEC	sp_configure	'clr enabled',	1
RECONFIGURE
go
----------
ALTER	DATABASE	damit	SET	TRUSTWORTHY	ON
GO
----------
CREATE	ASSEMBLY	IMAP4CLR
FROM	'C:\Program Files\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL\Binn\imap4clr.dll'
WITH	PERMISSION_SET=	UNSAFE
GO
----------
ALTER	ASSEMBLY	IMAP4CLR
FROM	'C:\Program Files\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL\Binn\imap4clr.dll';
GO
----------------------------------------------------------------------------------------------------
drop	FUNCTION	dbo.GetEMailsAvailable
go
create	FUNCTION	dbo.GetEMailsAvailable
(	@sServer	nvarchar ( 256 )
	,@iPort		int
	,@sUser		nvarchar ( 256 )
	,@sPassword	nvarchar ( 256 )
	,@dtStart	date=	null
)
/*
set	textsize		2147483647
set	quoted_identifier	on
set	arithabort		on		-- � ���������� ��������� ������ ������ ��� ��� ��������� ����
*/
RETURNS	TABLE
(	UID		int
	,MessageId	nvarchar ( 1000 )
	,Date		datetime
	,[From]		nvarchar ( 256 )
	,[To]		nvarchar ( max )
	,Cc		nvarchar ( max )
	,Subject	nvarchar ( 4000 )
	,Importance	nvarchar ( 6 )
)
EXTERNAL	NAME	IMAP4CLR.IMAP4CLR.GetEMailsAvailable
GO
----------------------------------------------------------------------------------------------------
drop	FUNCTION	dbo.GetEMails
go
create	FUNCTION	dbo.GetEMails
(	@sServer	nvarchar ( 256 )
	,@iPort		int
	,@sUser		nvarchar ( 256 )
	,@sPassword	nvarchar ( 256 )
	,@sUIDs		nvarchar ( 4000 )	-- max
)
/*
set	textsize		2147483647
set	quoted_identifier	on
set	arithabort		on		-- � ���������� ��������� ������ ������ ��� ��� ��������� ����
*/
RETURNS	TABLE
(	UID		int
	,MessageId	nvarchar ( 1000 )
	,Date		datetime
	,[From]		nvarchar ( 256 )
	,[To]		nvarchar ( max )
	,Cc		nvarchar ( max )
	,Subject	nvarchar ( 4000 )
	,Importance	nvarchar ( 6 )
	,Body		varbinary ( max ) )
EXTERNAL	NAME	IMAP4CLR.IMAP4CLR.GetEMails
GO
----------------------------------------------------------------------------------------------------
drop	proc	dbo.RemoveEMails
go
create	proc	dbo.RemoveEMails
(	@sServer	nvarchar ( 255 )
	,@iPort		int
	,@sUser		nvarchar ( 255 )
	,@sPassword	nvarchar ( 255 )
	,@sUIDs		nvarchar ( 4000 )	-- max
)
/*
set	textsize		2147483647
set	quoted_identifier	on
set	arithabort		on		-- � ���������� ��������� ������ ������ ��� ��� ��������� ����
*/
as
EXTERNAL	NAME	IMAP4CLR.IMAP4CLR.RemoveEMails
GO
----------------------------------------------------------------------------------------------------

/*
CLR:	/SEARCH ������ ���� uid � INBOX
	\�������� ������� uid,data,to,message-id						--uid �������� ��� ������������?
SQL:	/���������� ������� � �������� �� message-id ��� uid,data,to			--���� message-id �� �������� ��� �� ��������?
	\���������� ������ uid ������������ � �������������� � �������
CLR:	�� ����������� param=������ �������������� uid �������� � mail ������� ������ � ���� ������� � ������ ��������� �, ���� ����� bool param, �� ����� .eml
SQL:	��������� � �������
CLR:	�� ����������� param=������ ������������ uid ������� � mail �������
CLR+:	������� �������

select * from sys.dm_clr_loaded_assemblies
select * from sys.dm_clr_appdomains
select * from sys.dm_clr_properties
select * from sys.dm_clr_tasks

*/