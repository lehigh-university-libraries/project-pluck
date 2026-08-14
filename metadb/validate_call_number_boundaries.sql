--metadb:function validate_call_number_boundaries

DROP FUNCTION IF EXISTS validate_call_number_boundaries;

CREATE FUNCTION validate_call_number_boundaries(
    location_code TEXT,
    start_call_number_prefix TEXT,
    end_call_number_prefix TEXT
)
RETURNS TABLE (
    start_found BOOLEAN,
    end_found BOOLEAN
)
AS
$$
WITH
    ref_location AS (
        SELECT id FROM folio_inventory.location__t WHERE code = location_code
    ),
    start_boundary AS (
        SELECT COALESCE(item_notes.note, item.effective_shelving_order) AS shelving_order
        FROM folio_inventory.item__t item
        LEFT JOIN folio_inventory.holdings_record__t holdings ON item.holdings_record_id = holdings.id
        LEFT JOIN folio_derived.item_notes item_notes ON item_notes.item_id = item.id
            AND item_notes.note_type_name = 'Shelving order'
        WHERE item.effective_location_id = (SELECT id FROM ref_location)
          AND (item.item_level_call_number LIKE start_call_number_prefix || '%'
           OR holdings.call_number LIKE start_call_number_prefix || '%')
        LIMIT 1
    ),
    end_boundary AS (
        SELECT COALESCE(item_notes.note, item.effective_shelving_order) AS shelving_order
        FROM folio_inventory.item__t item
        LEFT JOIN folio_inventory.holdings_record__t holdings ON item.holdings_record_id = holdings.id
        LEFT JOIN folio_derived.item_notes item_notes ON item_notes.item_id = item.id
            AND item_notes.note_type_name = 'Shelving order'
        WHERE item.effective_location_id = (SELECT id FROM ref_location)
          AND (item.item_level_call_number LIKE end_call_number_prefix || '%'
           OR holdings.call_number LIKE end_call_number_prefix || '%')
        LIMIT 1
    )
SELECT
    EXISTS(SELECT 1 FROM start_boundary) AS start_found,
    EXISTS(SELECT 1 FROM end_boundary) AS end_found;
$$
LANGUAGE SQL;
