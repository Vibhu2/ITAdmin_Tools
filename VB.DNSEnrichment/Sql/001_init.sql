-- ============================================================
-- VB.DNSEnrichment -- Initial schema (migration 001)
-- ============================================================

CREATE TABLE IF NOT EXISTS Enrichment (
    IPAddress              TEXT PRIMARY KEY,
    Hostname               TEXT,
    HostnameSource         TEXT,
    MACAddress             TEXT,
    MACAddressNormalised   TEXT,
    Vendor                 TEXT,
    DeviceClass            TEXT,
    DeviceClassSource      TEXT,
    Confidence             TEXT,
    OSClass                TEXT,
    OperatingSystem        TEXT,
    Model                  TEXT,
    Location               TEXT,
    OU                     TEXT,
    OpenPorts              TEXT,
    HTTPTitle              TEXT,
    HTTPServer             TEXT,
    SNMPDescr              TEXT,
    RTSPBanner             TEXT,
    MDNSServiceType        TEXT,
    LeaseExpiry            TEXT,
    StepsAttempted         INTEGER,
    StepsSucceeded         INTEGER,
    StepsNoResult          INTEGER,
    StepsSkipped           INTEGER,
    StepsFailed            INTEGER,
    LayerTraceJson         TEXT,
    IsResolved             INTEGER,
    IsUnresolved           INTEGER,
    EnrichedAt             TEXT NOT NULL,
    EnrichmentDurationMs   INTEGER,
    FirstSeenAt            TEXT NOT NULL,
    UpdatedAt              TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS IX_Enrichment_DeviceClass ON Enrichment(DeviceClass);
CREATE INDEX IF NOT EXISTS IX_Enrichment_Hostname    ON Enrichment(Hostname);
CREATE INDEX IF NOT EXISTS IX_Enrichment_MAC         ON Enrichment(MACAddressNormalised);

CREATE TABLE IF NOT EXISTS EnrichmentHistory (
    Id                     INTEGER PRIMARY KEY AUTOINCREMENT,
    IPAddress              TEXT NOT NULL,
    OldHostname            TEXT,
    NewHostname            TEXT,
    OldMACAddress          TEXT,
    NewMACAddress          TEXT,
    OldDeviceClass         TEXT,
    NewDeviceClass         TEXT,
    ChangeReason           TEXT,
    ChangedAt              TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS IX_History_IP ON EnrichmentHistory(IPAddress);

CREATE TABLE IF NOT EXISTS SchemaVersion (
    Version    INTEGER PRIMARY KEY,
    AppliedAt  TEXT NOT NULL
);
