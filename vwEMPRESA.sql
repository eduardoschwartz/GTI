USE [Tributacao]
GO
/****** Object:  View [dbo].[vwFULLEMPRESA3]    Script Date: 07/27/2010 09:39:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[vwFULLEMPRESA3]
AS
SELECT DISTINCT 
                         codigomob, razaosocial, inscestadual, cnpj, CONVERT(VARCHAR(10), CASE WHEN abrevtipolog IS NULL THEN '' ELSE RTRIM(abrevtipolog) END) 
                         + ' ' + CONVERT(VARCHAR(10), CASE WHEN abrevtitlog IS NULL THEN '' ELSE RTRIM(abrevtitlog) END) + ' ' + CASE WHEN NOMELOGRADOURO IS NULL 
                         THEN NOMELOGR ELSE NOMELOGRADOURO END AS LOGRADOURO, numero, CONVERT(CHAR(5), codatividade) + '-' + descatividade AS ATIVIDADE, 
                         CASE SIMPLES WHEN 0 THEN 'N' WHEN 1 THEN 'S' END AS SIMPLES, CASE WHEN DATAENCERRAMENTO IS NULL THEN 'N' ELSE 'S' END AS ENCERRADO, 
                         dataabertura, dataencerramento, descbairro, desccidade, descuf, codlogradouro, complemento, respcontabil, alvara, cpf, cep, siglauf, ativextenso, deschorario, 
                         isentotaxa, mei, fonecontato, nomeesc, codatividade, vistoria, areatl, capitalsocial, qtdeempregado, orgao, rg, nomeorgao, numregistroresp, numprocencerramento, 
                         ruaesc, numeroesc, nomebairro, cepesc, nomecidade, uf, telefone, nomecontato, numprocesso, nomefantasia, horario, dataprocesso, dataprocencerramento, 
                         isencao, codigoaliq, descatividade, homepage, cargocontato, faxcontato, emailcontato, EECODLOGR, EENOMELOGR, EENUMERO, EECOMPL, EEUF, EECODCIDADE, 
                         EECODBAIRRO, EECEP, EEDESCBAIRRO, EEDESCCIDADE, qtdeprof, NOMELOGR, tipo, NOMERESPONSAVEL, regespecial, email, horarioext, isseletro
FROM            dbo.vwCNSMOBILIARIO

GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[56] 4[31] 2[6] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1[50] 2[25] 3) )"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4[30] 2[40] 3) )"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2[38] 3) )"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4[50] 3) )"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 6
   End
   Begin DiagramPane = 
      PaneHidden = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "vwCNSMOBILIARIO"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 199
               Right = 229
            End
            DisplayFlags = 280
            TopColumn = 54
         End
      End
   End
   Begin SQLPane = 
      PaneHidden = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 78
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
   ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwFULLEMPRESA3'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'      Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 2940
         Alias = 1335
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwFULLEMPRESA3'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwFULLEMPRESA3'