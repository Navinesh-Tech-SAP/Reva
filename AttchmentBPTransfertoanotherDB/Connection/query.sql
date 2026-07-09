SELECT 
   T0."CardCode" ,
   
  '\\172.21.0.45\SAP-Attachment\'  as 'path' ,T2."FileName" + '.' +T2.FileExt
 
FROM [LIVE_REVA_University].dbo.OCRD T0
inner JOIN [LIVE_REVA_University].dbo.OATC T1
    ON T0."AtcEntry" = T1."AbsEntry"
inner JOIN [LIVE_REVA_University].dbo.ATC1 T2
    ON T1."AbsEntry" = T2."AbsEntry"
	inner join [REVA_LIVE].dbo.OCRD T3 on T3.CardCode = T0.CardCode  AND T3."CardName" = T0."CardName"  


ORDER BY T0."CardCode" ;
