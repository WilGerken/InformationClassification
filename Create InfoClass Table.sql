
CREATE TABLE [dbo].[InfoClass](
	[FileImported] [nvarchar](255) NULL,
	[DateImported] [datetime] NULL,
	[Action] [nvarchar](255) NULL,
	[FileFolder] [nvarchar](2000) NULL,
	[NTID] [nvarchar](255) NULL,
	[DateFlagged] [nvarchar](255) NULL,
	[AppVersion] [nvarchar](255) NULL,
	[ObjectID] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO