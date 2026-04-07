--metadb:function get_items_between_call_number_prefixes

DROP FUNCTION IF EXISTS get_items_between_call_number_prefixes;

CREATE FUNCTION get_items_between_call_number_prefixes(
    location_code TEXT,
    start_call_number_prefix TEXT,
    end_call_number_prefix TEXT,
    query_limit INTEGER,
    query_offset INTEGER
)
RETURNS TABLE (
    barcode TEXT,
    id UUID,
    effective_shelving_order TEXT,
    local_shelving_order TEXT,
    effective_call_number TEXT,
    item_status TEXT,
    material_type TEXT,
    statistical_codes TEXT,
    title TEXT,
    instance_uuid UUID,
    instance_hrid TEXT,
    contributor TEXT,
    publication_date TEXT,
    faculty_author BOOLEAN,
    legacy_circ_count INTEGER,
    folio_circ_count BIGINT,
    oclc_number TEXT,
    item_effective_location_name TEXT,
    holdings_permanent_location_name TEXT,
    damage_inventory_note TEXT,
    electronic_holdings TEXT
)
AS
$$
WITH
    -- 1. Resolve the location code to an ID
    ref_location AS (
        SELECT id
        FROM folio_inventory.location__t
        WHERE code = location_code
    ),
    -- 2. Find the boundaries within the location
    start_boundary AS (
        SELECT
            COALESCE(item_notes.note, item.effective_shelving_order) AS shelving_order,
            item.effective_location_id
        FROM folio_inventory.item__t item
        LEFT JOIN folio_inventory.holdings_record__t holdings ON item.holdings_record_id = holdings.id
        LEFT JOIN folio_derived.item_notes item_notes ON item_notes.item_id = item.id
            AND item_notes.note_type_name = 'Shelving order'
        WHERE item.effective_location_id = (SELECT id FROM ref_location)
          AND (item.item_level_call_number LIKE start_call_number_prefix || '%'
           OR holdings.call_number LIKE start_call_number_prefix || '%')
        ORDER BY COALESCE(item_notes.note, item.effective_shelving_order) COLLATE ucs_basic ASC
        LIMIT 1
    ),
    end_boundary AS (
        SELECT
            COALESCE(item_notes.note, item.effective_shelving_order) AS shelving_order,
            item.effective_location_id
        FROM folio_inventory.item__t item
        LEFT JOIN folio_inventory.holdings_record__t holdings ON item.holdings_record_id = holdings.id
        LEFT JOIN folio_derived.item_notes item_notes ON item_notes.item_id = item.id
            AND item_notes.note_type_name = 'Shelving order'
        WHERE item.effective_location_id = (SELECT id FROM ref_location)
          AND (item.item_level_call_number LIKE end_call_number_prefix || '%'
           OR holdings.call_number LIKE end_call_number_prefix || '%')
        ORDER BY COALESCE(item_notes.note, item.effective_shelving_order) COLLATE ucs_basic DESC
        LIMIT 1
    ),
    -- 3. Identify just the items and instances in the range (NARROW)
    filtered_range AS (
        SELECT 
            item.id AS item_id,
            holdings.instance_id,
            item.barcode,
            item.effective_shelving_order,
            item_notes.note AS local_shelving_order,
            item.effective_location_id,
            item.material_type_id,
            holdings.permanent_location_id AS holdings_permanent_location_id,
            item.discovery_suppress
        FROM folio_inventory.item__t item
        LEFT JOIN folio_inventory.holdings_record__t holdings ON item.holdings_record_id = holdings.id
        LEFT JOIN folio_derived.item_notes item_notes ON item_notes.item_id = item.id 
            AND item_notes.note_type_name = 'Shelving order'
        WHERE
            item.effective_location_id = (SELECT id FROM ref_location)
            AND COALESCE(item_notes.note, item.effective_shelving_order) >=
                (SELECT shelving_order FROM start_boundary) COLLATE ucs_basic
            AND COALESCE(item_notes.note, item.effective_shelving_order) <= 
                (SELECT shelving_order FROM end_boundary) COLLATE ucs_basic
            AND (item.discovery_suppress IS NULL OR NOT item.discovery_suppress)
            AND item.barcode IS NOT NULL
        ORDER BY COALESCE(item_notes.note, item.effective_shelving_order) COLLATE ucs_basic
        LIMIT query_limit OFFSET query_offset
    ),
    -- 4. The "Many-to-One" data (Summaries and Counts)
    summarized_contributors AS (
        SELECT instance_id, STRING_AGG(contributor_name, '; ') as names
        FROM folio_derived.instance_contributors
        WHERE instance_id IN (SELECT instance_id FROM filtered_range)
        GROUP BY instance_id
    ),
    summarized_publications AS (
        SELECT instance_id, STRING_AGG(date_of_publication, '; ') as dates
        FROM folio_derived.instance_publication
        WHERE instance_id IN (SELECT instance_id FROM filtered_range)
        GROUP BY instance_id
    ),
    summarized_oclc AS (
        SELECT 
            instance_id, 
            STRING_AGG(TRIM(REPLACE(REPLACE(identifier, '(OCoLC)', ''), 'ocn', '')), '; ') AS identifiers
        FROM folio_derived.instance_identifiers
        WHERE instance_id IN (SELECT instance_id FROM filtered_range)
          AND identifier_type_name = 'OCLC'
        GROUP BY instance_id
    ),
    counted_folio_circ AS (
        SELECT (jsonb->'loan'->>'itemId')::UUID AS item_id, COUNT(*) AS checkout_count
        FROM folio_circulation.audit_loan
        WHERE (jsonb->'loan'->>'itemId')::UUID IN (SELECT item_id FROM filtered_range)
          AND jsonb->'loan'->>'action' = 'checkedout'
        GROUP BY (jsonb->'loan'->>'itemId')::UUID
    ),
    summarized_statistical_codes AS (
        SELECT item_id, STRING_AGG(statistical_code, '; ') AS codes
        FROM folio_derived.item_statistical_codes
        WHERE item_id IN (SELECT item_id FROM filtered_range)
        GROUP BY item_id
    ),
    -- 5. The "One-to-One" data (Lookups/References)
    ref_faculty_status AS (
        SELECT DISTINCT instance_id, TRUE as is_faculty
        FROM folio_derived.instance_notes
        WHERE instance_id IN (SELECT instance_id FROM filtered_range)
          AND instance_note = 'Lehigh Faculty Author Publication'
    ),
    ref_legacy_circ AS (
        SELECT item_id, note
        FROM folio_derived.item_notes
        WHERE item_id IN (SELECT item_id FROM filtered_range)
          AND note_type_name = 'OLE-Circ-Count'
    ),
    ref_locations AS (
        SELECT id, name
        FROM folio_inventory.location__t
        WHERE id IN (SELECT effective_location_id FROM filtered_range)
           OR id IN (SELECT holdings_permanent_location_id FROM filtered_range)
    ),
    ref_material_types AS (
        SELECT id, name
        FROM folio_inventory.material_type__t
        WHERE id IN (SELECT material_type_id FROM filtered_range)
    ),
    ref_damage AS (
        SELECT DISTINCT ON (item_ext.item_id) item_ext.item_id, item_notes.note
        FROM folio_derived.item_ext
        JOIN folio_derived.item_notes ON item_notes.item_id = item_ext.item_id
            AND item_notes.note_type_name = 'Inventoried Condition'
        WHERE item_ext.item_id IN (SELECT item_id FROM filtered_range)
          AND item_ext.damaged_status_name = 'Damaged'
        ORDER BY item_ext.item_id
    ),
    electronic_holdings_data AS (
        SELECT
            hn_print.note AS instance_hrid,
            json_agg(json_build_object(
                'holdings_hrid', holdings.hrid,
                'access_method', hn_access.note,
                'provider', hn_provider.note
            ))::TEXT AS holdings_data
        FROM folio_derived.holdings_notes hn_print
        LEFT JOIN folio_inventory.holdings_record__t holdings ON holdings.id = hn_print.holding_id
        LEFT JOIN folio_derived.holdings_notes hn_access
            ON hn_access.holding_id = hn_print.holding_id
            AND hn_access.note_type_name = 'Ebook access method'
        LEFT JOIN folio_derived.holdings_notes hn_provider
            ON hn_provider.holding_id = hn_print.holding_id
            AND hn_provider.note_type_name = 'Ebook provider'
        WHERE hn_print.note_type_name = 'Print version (I-HRID)'
          AND hn_print.note IN (
              SELECT inst2.hrid
              FROM folio_inventory.instance__t inst2
              WHERE inst2.id IN (SELECT instance_id FROM filtered_range)
          )
        GROUP BY hn_print.note
    )
-- 6. Final Select
SELECT 
    filtered_range.barcode, 
    filtered_range.item_id as id,
    filtered_range.effective_shelving_order, 
    filtered_range.local_shelving_order, 
    jsonb_extract_path_text(item_raw.jsonb, 'effectiveCallNumberComponents', 'callNumber') AS effective_call_number,
    jsonb_extract_path_text(item_raw.jsonb, 'status', 'name') AS item_status,
    ref_material_types.name AS material_type,
    summarized_statistical_codes.codes AS statistical_codes,
    inst.title,
    inst.id as instance_uuid,
    inst.hrid as instance_hrid,
    summarized_contributors.names AS contributor,
    summarized_publications.dates AS publication_date,
    COALESCE(ref_faculty_status.is_faculty, FALSE) AS faculty_author,
    COALESCE(NULLIF(regexp_replace(ref_legacy_circ.note, '\D', '', 'g'), ''), '0')::INTEGER AS legacy_circ_count,
    COALESCE(counted_folio_circ.checkout_count, 0) AS folio_circ_count,
    summarized_oclc.identifiers AS oclc_number,
    loc_eff.name AS item_effective_location_name,
    loc_perm.name AS holdings_permanent_location_name,
    ref_damage.note AS damage_inventory_note,
    electronic_holdings_data.holdings_data AS electronic_holdings
FROM filtered_range
LEFT JOIN folio_inventory.item item_raw ON filtered_range.item_id = item_raw.id
LEFT JOIN folio_inventory.instance__t inst ON filtered_range.instance_id = inst.id
LEFT JOIN summarized_contributors ON filtered_range.instance_id = summarized_contributors.instance_id
LEFT JOIN summarized_publications ON filtered_range.instance_id = summarized_publications.instance_id
LEFT JOIN ref_faculty_status ON filtered_range.instance_id = ref_faculty_status.instance_id
LEFT JOIN ref_legacy_circ ON filtered_range.item_id = ref_legacy_circ.item_id
LEFT JOIN counted_folio_circ ON filtered_range.item_id = counted_folio_circ.item_id
LEFT JOIN summarized_oclc ON filtered_range.instance_id = summarized_oclc.instance_id
LEFT JOIN ref_locations AS loc_eff ON filtered_range.effective_location_id = loc_eff.id
LEFT JOIN ref_locations AS loc_perm ON filtered_range.holdings_permanent_location_id = loc_perm.id
LEFT JOIN ref_material_types ON filtered_range.material_type_id = ref_material_types.id
LEFT JOIN ref_damage ON filtered_range.item_id = ref_damage.item_id
LEFT JOIN summarized_statistical_codes ON filtered_range.item_id = summarized_statistical_codes.item_id
LEFT JOIN electronic_holdings_data ON inst.hrid = electronic_holdings_data.instance_hrid
ORDER BY COALESCE(filtered_range.local_shelving_order, filtered_range.effective_shelving_order) COLLATE ucs_basic;
$$
LANGUAGE SQL;
