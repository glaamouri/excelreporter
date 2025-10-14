/*
 Étape 1 : Créer une "carte" des dates pour lier chaque date à sa précédente.
 La fonction LAG(colonne, 1) OVER (ORDER BY colonne) récupère la valeur de la ligne précédente,
 en se basant sur l'ordre spécifié (ici, l'ordre chronologique des dates).
*/
WITH DateMap AS (
    SELECT
        DWH_DATE_REF,
        LAG(DWH_DATE_REF, 1) OVER (ORDER BY DWH_DATE_REF) AS PREVIOUS_DWH_DATE_REF
    FROM (
        -- On ne prend que les dates uniques pour construire la carte
        SELECT DISTINCT DWH_DATE_REF
        FROM "NCIBL"."OBSTILUI9_T"
    ) AS DistinctDates
),

/*
 Étape 2 : Préparer les données de base.
 C'est votre requête originale, légèrement réorganisée en CTE pour être réutilisable.
*/
BaseData AS (
    SELECT
        T1.DWH_DATE_REF, T1.STIRAC, T1.STIGRE, T1.STIRUB, T1.STIMON, T1.STIORG, T1.STIMNT2, T1.STIMNT4, T1.STITPO,
        T1.FRRGRPCTRP, T1.STITGR, T2.ACC_TYPE, T2.COMPTE, T2.SOUS_COMPTE,
        CASE
            WHEN ISNULL(T1.STIECH, 0) > 0 THEN CONVERT(DATETIME, CONVERT(VARCHAR(8), T1.STIECH), 112)
            ELSE -1
        END AS TOT_O,
        CASE
            WHEN ISNULL(T1.STIECH, 0) > 0 THEN
                CASE
                    WHEN CONVERT(DATETIME, CONVERT(VARCHAR(8), T1.STIDVA), 112) <= 183 THEN 'moins de 6 mois'
                    WHEN CONVERT(DATETIME, CONVERT(VARCHAR(8), T1.STIECH), 112) <= 365 THEN 'moins d''un an'
                    ELSE 'plus d''un an'
                END
            ELSE ' '
        END AS BUCKET
    FROM "NCIBL"."OBSTILUI9_T" AS T1
    LEFT JOIN "NCIBL"."WKFS.LKP_IFRSLUX" AS T2 ON T2.PLNB = T1.STIPLNB
    WHERE
        -- Conservez vos filtres d'origine
        T1.DWH_DATE_REF >= DATEADD(month, -1, CAST(GETDATE() AS DATE))
        AND T1.STITPO = 'P'
        AND T1.STIIGR = '0'
        AND (T1.STIMNT2 <> 0 OR T1.STIMNT4 <> 0)
        AND T1.STIGRE < 800
)

/*
 Étape 3 : Sélection finale qui assemble les données du jour actuel et du jour précédent.
*/
SELECT
    T_Actuel.*, -- Sélectionne toutes les colonnes de la date de référence actuelle

    -- Ajoute les colonnes de la date de référence précédente en les renommant
    T_Precedent.DWH_DATE_REF    AS DWH_DATE_REF_PREVIOUS,
    T_Precedent.STIMON          AS STIMON_PREVIOUS,
    T_Precedent.STIMNT2         AS STIMNT2_PREVIOUS,
    T_Precedent.STIMNT4         AS STIMNT4_PREVIOUS,
    T_Precedent.STIORG          AS STIORG_PREVIOUS,
    T_Precedent.TOT_O           AS TOT_O_PREVIOUS,
    T_Precedent.BUCKET          AS BUCKET_PREVIOUS
    -- Ajoutez ici toute autre colonne du jour précédent que vous souhaitez afficher

FROM BaseData AS T_Actuel

-- Jointure avec la carte des dates pour trouver la date précédente pour la ligne actuelle
LEFT JOIN DateMap ON T_Actuel.DWH_DATE_REF = DateMap.DWH_DATE_REF

-- Jointure sur les données de base une seconde fois pour récupérer les informations de la date précédente
LEFT JOIN BaseData AS T_Precedent
    ON DateMap.PREVIOUS_DWH_DATE_REF = T_Precedent.DWH_DATE_REF
    /*
     !!! POINT CRUCIAL !!!
     La condition de jointure ci-dessous doit se baser sur la clé qui identifie
     un élément unique à travers le temps. J'ai supposé la combinaison suivante.
     Veuillez la vérifier et l'adapter à votre modèle de données.
    */
    AND T_Actuel.STIRAC = T_Precedent.STIRAC
    AND T_Actuel.STIRUB = T_Precedent.STIRUB
    AND T_Actuel.COMPTE = T_Precedent.COMPTE
    AND T_Actuel.SOUS_COMPTE = T_Precedent.SOUS_COMPTE

WHERE
    -- Vous pouvez remettre un filtre sur les dates ici si vous ne voulez que certains jours
    T_Actuel.DWH_DATE_REF IN ('2025-09-26', '2025-09-25', '2025-09-24', '2025-09-22', '2025-09-19', '2025-09-17', '2025-09-16', '2025-09-15')

ORDER BY
    T_Actuel.DWH_DATE_REF DESC, -- Ordonner par date la plus récente en premier
    T_Actuel.STIRAC,
    T_Actuel.STIRUB;
