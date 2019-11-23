select
 	 [_Diarienummer_], 
    [tempExcel].[dbo].[NuÖppnaÄrende].[_kommentar_]
 	 
from 
  [tempExcel].[dbo].[NuÖppnaÄrende] 
  join (
    select 
      count (*) as antalH, 
      [_rendenummer] 
    FROM 
      [tempExcel].[dbo].[w32] as w3 
      left outer join [tempExcel].[dbo].[vision2] as vision on [Diarienummer] = [_rendenummer] 
      AND vision.[Hõndelsedatum] = w3.[Ankomstdatum] 
      AND w3.[Riktning] = (
        case when vision.[Riktning] = 'UtgÕende' then 'Ut' when vision.[Riktning] = 'Inkommande' then 'In' else Null end
      ) 
      AND w3.[_tgõrd_Handling] = vision.[Rubrik] 
    where 
      [Hõndelsedatum] is null 
    group by 
      _rendenummer
	
  ) as temp2 on temp2._rendenummer = ltrim(RTRIM(substring(
    [tempExcel].[dbo].[NuÖppnaÄrende].[_kommentar_], 
    14, 14
  )))

  Order by _kommentar_ desc