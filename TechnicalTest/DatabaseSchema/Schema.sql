USE [DatabaseName]
GO
/****** Object:  Table [dbo].[tblCSV]    Script Date: 6/2/2020 1:23:36 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblCSV](
	[Syskey] [uniqueidentifier] NOT NULL,
	[TransID] [nvarchar](50) NOT NULL,
	[TransDate] [datetime] NOT NULL,
	[RecordStatus] [nvarchar](50) NOT NULL,
 CONSTRAINT [PK_tblCSV_1] PRIMARY KEY CLUSTERED 
(
	[Syskey] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tblCSVPaymentDetail]    Script Date: 6/2/2020 1:23:36 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblCSVPaymentDetail](
	[Syskey] [uniqueidentifier] NOT NULL,
	[CSVId] [uniqueidentifier] NOT NULL,
	[Amount] [decimal](18, 2) NOT NULL,
	[CurrencyCode] [nvarchar](50) NOT NULL,
 CONSTRAINT [PK_tblCSVPaymentDetail] PRIMARY KEY CLUSTERED 
(
	[Syskey] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tblExcel]    Script Date: 6/2/2020 1:23:36 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblExcel](
	[Syskey] [uniqueidentifier] NOT NULL,
	[TransID] [nvarchar](50) NOT NULL,
	[TransDate] [datetime] NOT NULL,
	[RecordStatus] [nvarchar](50) NOT NULL,
 CONSTRAINT [PK_tblExcel_1] PRIMARY KEY CLUSTERED 
(
	[Syskey] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tblXMLPaymentDetail]    Script Date: 6/2/2020 1:23:36 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblXMLPaymentDetail](
	[Syskey] [uniqueidentifier] NOT NULL,
	[XMLId] [uniqueidentifier] NOT NULL,
	[Amount] [decimal](18, 2) NOT NULL,
	[CurrencyCode] [nvarchar](50) NOT NULL,
 CONSTRAINT [PK_tblXMLPaymentDetail] PRIMARY KEY CLUSTERED 
(
	[Syskey] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
