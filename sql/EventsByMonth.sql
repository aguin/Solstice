SELECT sag.eventMonth, sag.num_sag_events, swl.num_swl_events
FROM (SELECT strftime('%Y-%m',event_start) || '-01' as eventMonth, count(event_start) as num_sag_events
		FROM EVENTS
		WHERE EVENT_TYPE = 'SAG'
		AND METER_ID = :METERID
		GROUP BY strftime('%Y-%m',event_start)
		) sag
LEFT JOIN (SELECT strftime('%Y-%m',event_start) || '-01' as eventMonth, count(event_start) as num_swl_events
		FROM EVENTS
		WHERE EVENT_TYPE = 'SWL'
		AND METER_ID = :METERID
		GROUP BY strftime('%Y-%m',event_start)
		) swl ON swl.eventMonth = sag.eventMonth
ORDER BY 1 ASC
