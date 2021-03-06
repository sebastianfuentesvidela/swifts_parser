if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[T_ECX_SwiftMT103]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[T_ECX_SwiftMT103]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[T_ECX_SwiftMT202]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[T_ECX_SwiftMT202]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[T_ECX_SwiftMT910]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[T_ECX_SwiftMT910]
GO

CREATE TABLE [dbo].[T_ECX_SwiftMT103] (
	[F_id] [int] IDENTITY (1, 1) NOT NULL ,
	[F_ISN] [int] NULL ,
	[F_tipo_mensaje] [int] NULL ,
	[F_Hora_ingreso] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Corresponsal] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
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
	[F_benRut] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_benName] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_remitInfo] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_cargoTo] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_senderChargCurr] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_senderChargAmou] [money] NULL ,
	[F_recverChargCurr] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_recverChargAmou] [money] NULL ,
	[F_send2recvInfo] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_regulReport] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_SwiftMessage] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_DataTail] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[T_ECX_SwiftMT202] (
	[F_id] [int] IDENTITY (1, 1) NOT NULL ,
	[F_ISN] [int] NULL ,
	[F_tipo_mensaje] [int] NULL ,
	[F_Hora_ingreso] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Corresponsal] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
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
	[F_DataTail] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[T_ECX_SwiftMT910] (
	[F_id] [int] IDENTITY (1, 1) NOT NULL ,
	[F_ISN] [int] NULL ,
	[F_tipo_mensaje] [int] NULL ,
	[F_Hora_ingreso] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[F_Corresponsal] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
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
	[F_DataTail] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

