CREATE TABLE [dbo].[Subscription]
(
	[Id] INT NOT NULL PRIMARY KEY, 
    [GroupName] NVARCHAR(100) NOT NULL, 
    [WebHookUri] NVARCHAR(MAX) NOT NULL
)
