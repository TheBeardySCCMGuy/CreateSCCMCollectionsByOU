select distinct System_OU_Name0 
,right(replace(System_OU_Name0,'/computers',''),CHARINDEX('/',REVERSE(replace(System_OU_Name0,'/computers','')))-1) as Office
,case when  System_OU_Name0 like ('%Europe Middle East Africa%') then 'EMEA'
	  when System_OU_Name0 like ('%americas%') then 'Americas'
	  when System_OU_Name0 like ('%Asia Pacific%') then 'APAC'
	  else 'Unknown Region'
	  End Region
from System_System_OU_Name_ARR
where System_OU_Name0 like 'Production.MyDomain.COM%'
--and System_OU_Name0 like '%computers%'
order by Office