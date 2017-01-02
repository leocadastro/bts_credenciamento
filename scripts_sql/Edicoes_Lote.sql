USE [Credenciamento2012]
GO

/****** Object:  Table [dbo].[Edicoes_Lote]    Script Date: 02/01/2017 18:36:18 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Edicoes_Lote](
	[ID_Lote_Edicao] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
	[ID_Edicao] [numeric](18, 0) NOT NULL,
	[Data_Inicio] [datetime] NULL,
	[Data_Fim] [datetime] NULL,
	[Ativo] [bit] NULL,
	[Valor] [decimal](18, 0) NULL,
	[Data_Cadastro] [datetime] NULL,
 CONSTRAINT [PK_Edicoes_Lote] PRIMARY KEY CLUSTERED 
(
	[ID_Lote_Edicao] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[Edicoes_Lote]  WITH CHECK ADD  CONSTRAINT [FK_Edicoes_Lote_Eventos_Edicoes] FOREIGN KEY([ID_Edicao])
REFERENCES [dbo].[Eventos_Edicoes] ([ID_Edicao])
GO

ALTER TABLE [dbo].[Edicoes_Lote] CHECK CONSTRAINT [FK_Edicoes_Lote_Eventos_Edicoes]
GO


