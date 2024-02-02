USE [db_Belashev]
GO

/****** Object:  Table [dbo].[BooksReaders]    Script Date: 02.02.2024 11:15:39 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[BooksReaders](
	[ID] [int] IDENTITY NOT NULL,
	[ID_book] [int] NOT NULL,
	[ID_reader] [int] NOT NULL,
 CONSTRAINT [PK_BooksReaders] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[BooksReaders]  WITH CHECK ADD  CONSTRAINT [FK_BooksReaders_Books] FOREIGN KEY([ID_book])
REFERENCES [dbo].[Books] ([ID_book])
GO

ALTER TABLE [dbo].[BooksReaders] CHECK CONSTRAINT [FK_BooksReaders_Books]
GO

ALTER TABLE [dbo].[BooksReaders]  WITH CHECK ADD  CONSTRAINT [FK_BooksReaders_Readers] FOREIGN KEY([ID_reader])
REFERENCES [dbo].[Readers] ([ID_reader])
GO

ALTER TABLE [dbo].[BooksReaders] CHECK CONSTRAINT [FK_BooksReaders_Readers]
GO


