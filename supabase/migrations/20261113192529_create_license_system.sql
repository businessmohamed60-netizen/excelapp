/*
# Systeme de gestion de licences web pour IATF Consultant Assistant

1. New Tables
- `licenses` : stocke les licences generees (cle, edition, expiration, statut)
  - id (uuid PK)
  - license_key (text, unique) : cle au format IATF-XXXXX-XXXXX-XXXXX-XXXXX
  - edition (text) : STANDARD ou PREMIUM
  - machine_id (text, nullable) : empreinte machine du client
  - customer_name (text) : nom du client
  - customer_email (text) : email du client
  - expires_at (date) : date d'expiration
  - status (text) : ACTIVE, EXPIRED, REVOKED, PENDING
  - max_activations (int, default 1) : nombre max d'activations
  - activation_count (int, default 0) : nombre d'activations realisees
  - created_at (timestamptz)
  - updated_at (timestamptz)

- `license_activations` : historique des activations par machine
  - id (uuid PK)
  - license_id (uuid FK -> licenses)
  - machine_id (text) : empreinte machine
  - machine_info (text) : infos complementaires (COMPUTERNAME, USERNAME)
  - activated_at (timestamptz)
  - ip_address (text, nullable)

2. Security
- Enable RLS on both tables.
- Grants: anon + authenticated can verify a license (SELECT), only service_role can create/update/delete.
- License verification endpoint is public (anon can call).
*/

CREATE TABLE IF NOT EXISTS licenses (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    license_key text UNIQUE NOT NULL,
    edition text NOT NULL DEFAULT 'STANDARD',
    machine_id text,
    customer_name text,
    customer_email text,
    expires_at date NOT NULL,
    status text NOT NULL DEFAULT 'PENDING',
    max_activations int NOT NULL DEFAULT 1,
    activation_count int NOT NULL DEFAULT 0,
    created_at timestamptz DEFAULT now(),
    updated_at timestamptz DEFAULT now()
);

CREATE TABLE IF NOT EXISTS license_activations (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    license_id uuid NOT NULL REFERENCES licenses(id) ON DELETE CASCADE,
    machine_id text NOT NULL,
    machine_info text,
    activated_at timestamptz DEFAULT now(),
    ip_address text
);

ALTER TABLE licenses ENABLE ROW LEVEL SECURITY;
ALTER TABLE license_activations ENABLE ROW LEVEL SECURITY;

GRANT SELECT ON licenses TO anon, authenticated;
GRANT SELECT ON license_activations TO anon, authenticated;
GRANT INSERT, UPDATE, DELETE ON licenses TO service_role;
GRANT INSERT, UPDATE, DELETE ON license_activations TO service_role;

DROP POLICY IF EXISTS "anon_select_licenses" ON licenses;
CREATE POLICY "anon_select_licenses" ON licenses FOR SELECT
    TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_select_activations" ON license_activations;
CREATE POLICY "anon_select_activations" ON license_activations FOR SELECT
    TO anon, authenticated USING (true);

CREATE INDEX IF NOT EXISTS idx_licenses_key ON licenses(license_key);
CREATE INDEX IF NOT EXISTS idx_licenses_status ON licenses(status);
CREATE INDEX IF NOT EXISTS idx_licenses_email ON licenses(customer_email);
CREATE INDEX IF NOT EXISTS idx_activations_license_id ON license_activations(license_id);
CREATE INDEX IF NOT EXISTS idx_activations_machine_id ON license_activations(machine_id);

CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = now();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS licenses_updated_at ON licenses;
CREATE TRIGGER licenses_updated_at BEFORE UPDATE ON licenses
    FOR EACH ROW EXECUTE FUNCTION update_updated_at();
