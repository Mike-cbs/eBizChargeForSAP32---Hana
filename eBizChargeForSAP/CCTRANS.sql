
/****** Object:  Table [dbo].[CCTRANS]    Script Date: 11/12/2014 15:44:54 ******/
IF Not  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CCTRANS]') AND type in (N'U'))
CREATE TABLE [dbo].[CCTRANS](
	[customerID] [nvarchar](10) NULL,
	[customerName] [nvarchar](50) NULL,
	[crCardNum] [nvarchar](50) NULL,
	[Description] [nvarchar](50) NULL,
	[recID] [nvarchar](20) NULL,
	[acsUrl] [nvarchar](80) NULL,
	[authAmount] [nvarchar](50) NULL,
	[authCode] [nvarchar](10) NULL,
	[avsResult] [nvarchar](255) NULL,
	[avsResultCode] [nvarchar](10) NULL,
	[batchNum] [nvarchar](20) NULL,
	[batchRef] [nvarchar](20) NULL,
	[cardCodeResult] [nvarchar](255) NULL,
	[cardCodeResultCode] [nvarchar](10) NULL,
	[cardLevelResult] [nvarchar](255) NULL,
	[cardLevelResultCode] [nvarchar](10) NULL,
	[conversionRate] [nvarchar](20) NULL,
	[convertedAmount] [nvarchar](20) NULL,
	[convertedAmountCurrency] [nvarchar](20) NULL,
	[custNum] [nvarchar](50) NULL,
	[error] [nvarchar](50) NULL,
	[errorCode] [nvarchar](10) NULL,
	[isDuplicate] [nvarchar](10) NULL,
	[payload] [nvarchar](50) NULL,
	[profilerScore] [nvarchar](10) NULL,
	[profilerResponse] [nvarchar](50) NULL,
	[profilerReason] [nvarchar](50) NULL,
	[refNum] [nvarchar](10) NULL,
	[remainingBalance] [nvarchar](20) NULL,
	[result] [nvarchar](50) NULL,
	[resultCode] [nvarchar](10) NULL,
	[status] [nvarchar](50) NULL,
	[statusCode] [nvarchar](10) NULL,
	[vpasResultCode] [nvarchar](10) NULL,
	[recDate] [datetime] NULL,
	[command] [nvarchar](20) NULL,
	[amount] [nvarchar](50) NULL
) ON [PRIMARY]



