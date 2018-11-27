select 
	CASE WHEN Sort IS NOT NULL THEN 'キーネーム：' + CAST(a.Sort as VARCHAR(10))
	WHEN Name IS NOT NULL THEN 'デフォルト値：' + a.Name
	WHEN Label_ja IS NOT NULL THEN 'その他：' + a.Label_ja
	ELSE '-' END as snl
from HELI_IT_Doc as a

