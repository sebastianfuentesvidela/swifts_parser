if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spr_swf_GetBank]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spr_swf_GetBank]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spr_swf_mensaje]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spr_swf_mensaje]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spr_swf_message]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spr_swf_message]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[T_ECX_SwiftMT103]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[T_ECX_SwiftMT103]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[T_ECX_SwiftMT202]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[T_ECX_SwiftMT202]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[T_ECX_SwiftMT799]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[T_ECX_SwiftMT799]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[T_ECX_SwiftMT910]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[T_ECX_SwiftMT910]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[T_Swift_Address]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[T_Swift_Address]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[T_Swift_Titles]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[T_Swift_Titles]
GO

CREATE TABLE [dbo].[T_ECX_SwiftMT103] (
	[F_id] [int] IDENTITY (1, 1) NOT NULL ,
	[F_ISN] [int] NULL ,
	[F_FechaIngreso] [datetime] NULL ,
	[F_FechaProceso] [datetime] NOT NULL ,
	[F_tipo_mensaje] [int] NULL ,
	[F_Hora_ingreso] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Corresponsal] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_RutCliente] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_timeType] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_timeTime] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_timeOffs] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_msgReference] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_bankOperatCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_instructCodeCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_instructCodeRefr] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_transacType] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Valuta] [datetime] NULL ,
	[F_Moneda] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Monto] [money] NULL ,
	[F_instructCurr] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_instructAmount] [money] NULL ,
	[F_exchRate] [money] NULL ,
	[F_orderingCusAcc] [varchar] (34) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_orderingCusName] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_origBank] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_correspSender] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_correspReceiver] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_terceraInstit] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_intermediariaInstit] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_financialInstit] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_benAccount] [varchar] (34) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_benName] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_remitInfo] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_cargoTo] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_senderChargCurr] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_senderChargAmou] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_recverChargCurr] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_recverChargAmou] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_send2recvInfo] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_regulReport] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftMessage] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_DataTail] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_OrigFile] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftSender] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftReceiver] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftMUR] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftHead] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftText] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftTrailer] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[T_ECX_SwiftMT202] (
	[F_id] [int] IDENTITY (1, 1) NOT NULL ,
	[F_ISN] [int] NULL ,
	[F_FechaIngreso] [datetime] NULL ,
	[F_FechaProceso] [datetime] NULL ,
	[F_tipo_mensaje] [int] NULL ,
	[F_Hora_ingreso] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Corresponsal] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_RutCliente] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_timeType] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_timeTime] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_timeOffs] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_msgReference] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_relReference] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Valuta] [datetime] NULL ,
	[F_Moneda] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Monto] [money] NULL ,
	[F_origBank] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_correspSender] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_correspReceiver] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_intermediariaInstit] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_financialInstit] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_benefInstit] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_send2recvInfo] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_regulReport] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftMessage] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_DataTail] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_OrigFile] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftSender] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftReceiver] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftMUR] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftHead] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftText] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftTrailer] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[T_ECX_SwiftMT799] (
	[F_id] [int] IDENTITY (1, 1) NOT NULL ,
	[F_ISN] [int] NULL ,
	[F_FechaIngreso] [datetime] NULL ,
	[F_FechaProceso] [datetime] NULL ,
	[F_tipo_mensaje] [int] NULL ,
	[F_Hora_ingreso] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Corresponsal] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_RutCliente] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_msgReference] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_relReference] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Explanation] [varchar] (1800) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftMessage] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_DataTail] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_OrigFile] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftSender] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftReceiver] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftMUR] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftHead] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftText] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftTrailer] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[T_ECX_SwiftMT910] (
	[F_id] [int] IDENTITY (1, 1) NOT NULL ,
	[F_ISN] [int] NULL ,
	[F_FechaIngreso] [datetime] NULL ,
	[F_FechaProceso] [datetime] NULL ,
	[F_tipo_mensaje] [int] NULL ,
	[F_Hora_ingreso] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Corresponsal] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_RutCliente] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_msgReference] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_relReference] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_accountIdent] [varchar] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Valuta] [datetime] NULL ,
	[F_Moneda] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Monto] [money] NULL ,
	[F_orderingCusAcc] [varchar] (34) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_orderingCusName] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_origBank] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_intermediariaInstit] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_send2recvInfo] [varchar] (210) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_regulReport] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftMessage] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_DataTail] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_OrigFile] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftSender] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftReceiver] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftMUR] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftHead] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftText] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftTrailer] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[T_Swift_Address] (
	[Id] [int] IDENTITY (1, 1) NOT NULL ,
	[BankBic] [varchar] (11) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BankName] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BankAddr] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col002] [char] (11) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col003] [char] (105) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col004] [char] (70) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col005] [char] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col006] [char] (64) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col007] [char] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col008] [char] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col009] [char] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col010] [char] (70) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col011] [char] (105) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col012] [char] (70) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col013] [char] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col014] [char] (105) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Col015] [char] (77) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[T_Swift_Titles] (
	[F_Tipo] [int] NOT NULL ,
	[F_Row] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[F_Message] [varchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Mensaje] [varchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[T_ECX_SwiftMT103] WITH NOCHECK ADD 
	CONSTRAINT [PK_T_ECX_SwiftMT103] PRIMARY KEY  CLUSTERED 
	(
		[F_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[T_ECX_SwiftMT202] WITH NOCHECK ADD 
	CONSTRAINT [PK_T_ECX_SwiftMT202] PRIMARY KEY  CLUSTERED 
	(
		[F_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[T_ECX_SwiftMT799] WITH NOCHECK ADD 
	CONSTRAINT [PK_T_ECX_SwiftMT799] PRIMARY KEY  CLUSTERED 
	(
		[F_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[T_ECX_SwiftMT910] WITH NOCHECK ADD 
	CONSTRAINT [PK_T_ECX_SwiftMT910] PRIMARY KEY  CLUSTERED 
	(
		[F_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[T_Swift_Address] WITH NOCHECK ADD 
	CONSTRAINT [PK_T_Swift_Address] PRIMARY KEY  CLUSTERED 
	(
		[Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[T_Swift_Titles] WITH NOCHECK ADD 
	CONSTRAINT [PK_T_Swift_Titles] PRIMARY KEY  CLUSTERED 
	(
		[F_Tipo],
		[F_Row]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[T_ECX_SwiftMT103] WITH NOCHECK ADD 
	CONSTRAINT [DF_T_ECX_SwiftMT103_F_FechaProceso] DEFAULT (getdate()) FOR [F_FechaProceso]
GO

ALTER TABLE [dbo].[T_ECX_SwiftMT202] WITH NOCHECK ADD 
	CONSTRAINT [DF_T_ECX_SwiftMT202_F_FechaProceso] DEFAULT (getdate()) FOR [F_FechaProceso]
GO

ALTER TABLE [dbo].[T_ECX_SwiftMT799] WITH NOCHECK ADD 
	CONSTRAINT [DF_T_ECX_SwiftMT799_F_FechaProceso] DEFAULT (getdate()) FOR [F_FechaProceso]
GO

ALTER TABLE [dbo].[T_ECX_SwiftMT910] WITH NOCHECK ADD 
	CONSTRAINT [DF_T_ECX_SwiftMT910_F_FechaProceso] DEFAULT (getdate()) FOR [F_FechaProceso]
GO

 CREATE  INDEX [IX_T_Swift_Address] ON [dbo].[T_Swift_Address]([BankBic]) ON [PRIMARY]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE spr_swf_GetBank (
	@bic varchar(11) )
AS
if exists(SELECT  1 FROM         T_Swift_Address  WHERE  BankBic = @bic)
	SELECT  Top 1  BankName, BankAddr
		FROM         T_Swift_Address
		WHERE     (BankBic = @bic)
Else
	if exists(SELECT  1 FROM         T_Swift_Address  WHERE  BankBic  like  left(@bic, 10) + '%')
	SELECT  TOP 1   BankName, BankAddr
		FROM         T_Swift_Address
		WHERE     (BankBic like  left(@bic, 10) + '%') 
	Else
		if exists(SELECT  1 FROM         T_Swift_Address  WHERE  BankBic  like  left(@bic, 9) + '%')
		SELECT  TOP 1   BankName, BankAddr
			FROM         T_Swift_Address
			WHERE     (BankBic like  left(@bic, 9) + '%') 
		Else
			if exists(SELECT  1 FROM         T_Swift_Address  WHERE  BankBic  like  left(@bic, 8) + '%')
			SELECT  TOP 1   BankName, BankAddr
				FROM         T_Swift_Address
				WHERE     (BankBic like  left(@bic, 8) + '%') 
			Else
				SELECT  '' AS BankName, '' AS BankAddr
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE spr_swf_mensaje
( @swftipo int, @swftag varchar(10) )
AS
if exists(SELECT  1 FROM         T_Swift_Titles WHERE     (F_Tipo = @swftipo) AND (F_Row = @swftag))
	SELECT     F_Mensaje
	FROM         T_Swift_Titles
	WHERE     (F_Tipo = @swftipo) AND (F_Row = @swftag)
ELSE
	SELECT     'tag indefinido' as F_Mensaje
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE spr_swf_message
( @swftipo int, @swftag varchar(10) )
AS
if exists(SELECT  1 FROM         T_Swift_Titles WHERE     (F_Tipo = @swftipo) AND (F_Row = @swftag))
	SELECT     F_Message
	FROM         T_Swift_Titles
	WHERE     (F_Tipo = @swftipo) AND (F_Row = @swftag)
ELSE
	SELECT     'tag indefinido' as F_Message
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

