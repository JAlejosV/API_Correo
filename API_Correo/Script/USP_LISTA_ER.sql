use InterfacesDB
go

IF OBJECT_ID('USP_LISTA_ER', 'P') IS NOT NULL
    DROP PROCEDURE USP_LISTA_ER
GO
  
CREATE PROC USP_LISTA_ER  
AS  
BEGIN  
 SELECT   
 CodigoEmpresa,  
 ComprobanteTesoreria,  
 Moneda,  
 CorrelativoHelm,                                      
 Fecha,                         
 TipoCambio  
 FROM EntregaRendir  
END  
  