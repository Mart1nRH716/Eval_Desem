-- Antes de hacelro en el 103, las pruebas se realizarán en local
USE [DBTRABAJODAS]
GO

CREATE TABLE eval_UMAE(
	id_eval_umae INT IDENTITY(0,1) NOT NULL,
	anio INT,
	mes TINYINT,
	cve_pre VARCHAR(12),
	nom_umae VARCHAR(40),
	-- color VARCHAR(10),
	indicador VARCHAR(40) NOT NULL,
	valor FLOAT
	CONSTRAINT [PK_eval_umae] PRIMARY KEY CLUSTERED
	(
		id_eval_UMAE ASC
	)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]


/*
INSERT INTO [11.254.16.103].[DBTRABAJODAS].[eval_UMAE]
SELECT * FROM eval_UMAE
where  anio = 2025 and mes = 2
--TRUNCATE TABLE eval_UMAE*/