// FOLIO connection
// Configure your institution's FOLIO server URLs and tenant ID.
const FOLIO_API_URL = {
  'prod': 'https://lehigh-okapi.folio.indexdata.com',
  'test': 'https://lehigh-test-okapi.folio.indexdata.com'
};
const TENANT = "lu";

// FOLIO item note types
// Names must match the note types configured in your FOLIO instance.
const DECISION_NOTE_ITEM_TYPE = 'Project Pluck Decision';
const MISSING_CHECK_IN_NOTE_TYPE = 'Check in';
const MISSING_CHECK_IN_NOTE_TEXT = 'Withdrawn.  Route to Cataloging.';
// Used only if the Faculty Author column is enabled.
const FACULTY_AUTHOR_NOTE_TEXT = "Lehigh Faculty Author Publication";
// Used only if the Legacy Circ Count column is enabled.
const LEGACY_CIRC_COUNT_NOTE_TYPE_NAME = 'OLE-Circ-Count';

// FOLIO identifiers
// Name must match the identifier type configured in your FOLIO instance.
const OCLC_NUMBER_IDENTIFIER_TYPE_NAME = 'OCLC';

// FOLIO instance statuses
// Name must match the instance status configured in your FOLIO instance.
const INSTANCE_STATUS_WITHDRAWN_CODE = 'Withdrawn';

// FOLIO statistical codes
// Names must match the statistical codes configured in your FOLIO instance.
const FINAL_STATE_KEEP = 'decision-keep-2024';
const FINAL_STATE_WITHDRAW = 'decision-withdraw-2024';
// Used only if the Retention column is enabled.
const RETENTION_CODES = [
  'EAST',
];
// Used only if the Inventoried column is enabled.
const INVENTORIED_CODES = [
  'INV-2025',
];

// OCLC consortium symbols
// List the OCLC symbols for each consortium your institution belongs to.
// Used to approximate consortium holding counts from the WorldCat API.
// Used only if the PALCI Holdings column is enabled.
const PALCI_OCLC_SYMBOLS = ['AVL','BEA','BMC','PBU','PBE','CRC','PMC','HHC','PBB','LQS','MAN','XR4','ALL','DKC','DRU','DXU','DUQ','ETS','EAS','ELZ','LFM','PGU','GDC','GBL','HUSAT','HVC','HFC','PZI','PJU','KOL','KZS','LRC','LAS','LAF','VFL','LVC','LYU','LYC','WHV','MRW','QRA','PGM','MVS','CMZ','NJM','MOR','EVI','ZMU','ZYU','UPM','CSC','REC','EIB','PHU','PMN','PTP','ROB','NJG','NJR','PSF','SJD','STH','SQP','SRS','PHA','SUS','PSC','TEU','PCT','PAU','PIT','SRU','URS','PVU','PUG','QWC','WVX','WVU','WFN','UWC','YCP'];
// Used only if the LVAIC Holdings column is enabled.
const LVAIC_OCLC_SYMBOLS = ['CC#', 'LAF', 'LYU', 'MOR', 'EVI', 'ALL'];

// Decisions
// The list of possible retention decisions shown in the spreadsheet drop-down.
// If you add or remove decisions, update DECISION_TO_FINAL_STATE below to match.
const WITHDRAW = 'Withdrawn';
const REMOTE = 'Move to remote storage';
const SC = 'Move to special collections';
const MISSING = 'Missing';
const NO_CHANGE = 'No change';
const DECISIONS = [WITHDRAW, REMOTE, SC, MISSING, NO_CHANGE];

// Maps each decision to its FOLIO final-state statistical code.
// Must include an entry for every decision listed in DECISIONS above.
const DECISION_TO_FINAL_STATE = new Map([
  [WITHDRAW, FINAL_STATE_WITHDRAW],
  [REMOTE,   FINAL_STATE_KEEP],
  [SC,       FINAL_STATE_KEEP],
  [MISSING,  FINAL_STATE_WITHDRAW],
  [NO_CHANGE,FINAL_STATE_KEEP],
]);
