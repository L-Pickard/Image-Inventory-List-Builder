SELECT iv.[Entity]
    ,CURRENT_TIMESTAMP AS [Last Refresh]
    ,iv.[Item No]
    ,CASE 
        WHEN iv.[Entity] = 'Shiner Ltd'
            THEN it.[Ltd GBP Unit Cost]
        WHEN iv.[Entity] = 'Shiner B.V'
            THEN it.[B.V EUR Unit Cost]
        WHEN iv.[Entity] = 'Shiner LLC'
            THEN it.[LLC USD Unit Cost]
        END AS [Unit Cost]
    ,br.[Brand Name]
    ,it.[Description]
    ,it.[Description 2]
    ,CASE 
        WHEN TRIM(it.[Size 1]) = ''
            THEN 'O/S'
        ELSE TRIM(it.[Size 1])
        END AS [Size 1]
    ,CASE 
        WHEN TRIM(it.[EU Size]) <> ''
            THEN TRIM(it.[EU Size])
        ELSE (
                CASE 
                    WHEN TRIM(it.[Size 1]) = ''
                        THEN 'O/S'
                    ELSE TRIM(it.[Size 1])
                    END
                )
        END AS [EU Size]
    ,CASE 
        WHEN TRIM(it.[US Size]) <> ''
            THEN TRIM(it.[US Size])
        ELSE (
                CASE 
                    WHEN TRIM(it.[Size 1]) = ''
                        THEN 'O/S'
                    ELSE TRIM(it.[Size 1])
                    END
                )
        END AS [US Size]
    ,it.[Size 1 Unit]
    ,it.[Colours]
    ,it.[Category Code] AS [Category]
    ,it.[Group Code] AS [Group]
    ,it.[Season]
    ,it.[Item Info]
    ,CASE 
        WHEN it.[On Sale] = 1
            THEN CHAR(252)
        ELSE CHAR(251)
        END AS [On Sale]
    ,it.[GBP Trade] AS [£ Price]
    ,it.[GBP SRP] AS [£ SRP]
    ,it.[EUR Trade] AS [€ Price]
    ,it.[EUR SRP] AS [€ SRP]
    ,it.[USD Trade] AS [$ Price]
    ,it.[USD SRP] AS [$ SRP]
    ,iv.[Free Stock]
    ,it.[EAN Barcode]
    ,it.[Tariff No]
    ,ISNULL(CO.[Country Name], 'Not Assigned') AS [Country of Origin]
    ,(
        SELECT TOP 1 TRIM([File Path])
        FROM [dThumbnail]
        WHERE [Item No] = it.[Item No]
        ) AS [URL]
FROM [dInventory] AS iv
LEFT JOIN [dItem] AS it
    ON iv.[Item No] = it.[Item No]
LEFT JOIN [dBrand] AS br
    ON it.[Brand Code] = br.[Brand Code]
LEFT JOIN [dCountry] AS co
    ON it.[COO] = co.[Country Code]
