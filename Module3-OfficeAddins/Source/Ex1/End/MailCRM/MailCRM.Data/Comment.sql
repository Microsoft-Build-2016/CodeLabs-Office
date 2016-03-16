CREATE TABLE [dbo].[Comment]
(
	[Id] INT NOT NULL PRIMARY KEY IDENTITY, 
    [ContactEmail] NCHAR(100) NOT NULL, 
    [CommentText] NVARCHAR(140) NOT NULL, 
    [PostDate] DATETIME NOT NULL
)
