import { useState, useEffect, useCallback } from "react";

interface License {
  id: string;
  license_key: string;
  edition: string;
  customer_name: string | null;
  customer_email: string | null;
  expires_at: string;
  status: string;
  max_activations: number;
  activation_count: number;
  created_at: string;
}

interface Stats {
  total: number;
  active: number;
  pending: number;
  expired: number;
  revoked: number;
  premium: number;
  standard: number;
  total_activations: number;
}

const API_URL = `${import.meta.env.VITE_SUPABASE_URL}/functions/v1/license-api`;
const API_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;

async function api(path: string, options: RequestInit = {}): Promise<any> {
  const res = await fetch(`${API_URL}${path}`, {
    ...options,
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${API_KEY}`,
      "X-Client-Info": "iatf-license-manager",
      ...options.headers,
    },
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Request failed");
  return data;
}

export default function App() {
  const [licenses, setLicenses] = useState<License[]>([]);
  const [stats, setStats] = useState<Stats | null>(null);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [showModal, setShowModal] = useState(false);
  const [generatedKey, setGeneratedKey] = useState<string | null>(null);
  const [toast, setToast] = useState<{ msg: string; type: "success" | "error" } | null>(null);
  const [page, setPage] = useState(1);
  const [total, setTotal] = useState(0);
  const perPage = 50;

  const showToast = (msg: string, type: "success" | "error" = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3000);
  };

  const loadLicenses = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page: String(page), per_page: String(perPage) });
      if (search) params.set("search", search);
      if (statusFilter) params.set("status", statusFilter);
      const data = await api(`/list?${params}`);
      setLicenses(data.licenses);
      setTotal(data.total);
    } catch (e: any) {
      showToast(e.message, "error");
    } finally {
      setLoading(false);
    }
  }, [page, search, statusFilter]);

  const loadStats = useCallback(async () => {
    try {
      const data = await api("/stats");
      setStats(data.stats);
    } catch (e) {
      // silent
    }
  }, []);

  useEffect(() => { loadLicenses(); }, [loadLicenses]);
  useEffect(() => { loadStats(); }, [loadStats]);

  const handleGenerate = async (form: any) => {
    try {
      const data = await api("/generate", {
        method: "POST",
        body: JSON.stringify(form),
      });
      setGeneratedKey(data.license.license_key);
      showToast("Licence generee avec succes");
      loadLicenses();
      loadStats();
    } catch (e: any) {
      showToast(e.message, "error");
    }
  };

  const handleRevoke = async (id: string) => {
    if (!confirm("Revoquer cette licence ?")) return;
    try {
      await api(`/revoke/${id}`, { method: "PUT" });
      showToast("Licence revoquee");
      loadLicenses();
      loadStats();
    } catch (e: any) {
      showToast(e.message, "error");
    }
  };

  const handleDelete = async (id: string) => {
    if (!confirm("Supprimer definitivement cette licence ?")) return;
    try {
      await api(`/${id}`, { method: "DELETE" });
      showToast("Licence supprimee");
      loadLicenses();
      loadStats();
    } catch (e: any) {
      showToast(e.message, "error");
    }
  };

  const handleExtend = async (id: string) => {
    const date = prompt("Nouvelle date d'expiration (AAAA-MM-JJ) :");
    if (!date) return;
    try {
      await api(`/extend/${id}`, { method: "PUT", body: JSON.stringify({ expires_at: date }) });
      showToast("Licence prolongee");
      loadLicenses();
      loadStats();
    } catch (e: any) {
      showToast(e.message, "error");
    }
  };

  const copyKey = (key: string) => {
    navigator.clipboard.writeText(key);
    showToast("Cle copiee");
  };

  const totalPages = Math.ceil(total / perPage);

  return (
    <div className="app">
      <div className="header">
        <div>
          <h1>IATF <span>License Manager</span></h1>
          <div className="subtitle">Gestion des licences — IATF Consultant Assistant</div>
        </div>
        <button className="btn btn-primary" onClick={() => { setGeneratedKey(null); setShowModal(true); }}>
          + Nouvelle licence
        </button>
      </div>

      {stats && (
        <div className="stats-grid">
          <div className="stat-card">
            <div className="label">Total</div>
            <div className="value">{stats.total}</div>
          </div>
          <div className="stat-card">
            <div className="label">Actives</div>
            <div className="value success">{stats.active}</div>
          </div>
          <div className="stat-card">
            <div className="label">En attente</div>
            <div className="value warning">{stats.pending}</div>
          </div>
          <div className="stat-card">
            <div className="label">Expirees</div>
            <div className="value error">{stats.expired}</div>
          </div>
          <div className="stat-card">
            <div className="label">Premium</div>
            <div className="value premium">{stats.premium}</div>
          </div>
          <div className="stat-card">
            <div className="label">Activations</div>
            <div className="value">{stats.total_activations}</div>
          </div>
        </div>
      )}

      <div className="toolbar">
        <input
          className="search-input"
          placeholder="Rechercher (nom, email, cle)..."
          value={search}
          onChange={(e) => { setSearch(e.target.value); setPage(1); }}
        />
        <select className="filter-select" value={statusFilter} onChange={(e) => { setStatusFilter(e.target.value); setPage(1); }}>
          <option value="">Tous les statuts</option>
          <option value="ACTIVE">Actives</option>
          <option value="PENDING">En attente</option>
          <option value="EXPIRED">Expirées</option>
          <option value="REVOKED">Révoquées</option>
        </select>
      </div>

      {loading ? (
        <div className="loading">Chargement...</div>
      ) : licenses.length === 0 ? (
        <div className="table-container">
          <div className="empty">
            <p>Aucune licence trouvée. Cliquez sur "Nouvelle licence" pour commencer.</p>
          </div>
        </div>
      ) : (
        <div className="table-container">
          <table>
            <thead>
              <tr>
                <th>Clé</th>
                <th>Client</th>
                <th className="hide-mobile">Édition</th>
                <th className="hide-mobile">Expiration</th>
                <th>Statut</th>
                <th className="hide-mobile">Activations</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {licenses.map((lic) => (
                <tr key={lic.id}>
                  <td>
                    <span className="license-key" style={{ cursor: "pointer" }} onClick={() => copyKey(lic.license_key)}>
                      {lic.license_key.substring(0, 16)}...
                    </span>
                  </td>
                  <td>
                    <div>{lic.customer_name || "—"}</div>
                    {lic.customer_email && (
                      <div style={{ fontSize: 12, color: "var(--text-dim)" }}>{lic.customer_email}</div>
                    )}
                  </td>
                  <td className="hide-mobile">
                    <span className={`badge ${lic.edition === "PREMIUM" ? "badge-premium" : "badge-standard"}`}>
                      {lic.edition}
                    </span>
                  </td>
                  <td className="hide-mobile" style={{ fontSize: 13 }}>{lic.expires_at}</td>
                  <td>
                    <span className={`badge badge-${lic.status.toLowerCase()}`}>{lic.status}</span>
                  </td>
                  <td className="hide-mobile">{lic.activation_count}/{lic.max_activations}</td>
                  <td>
                    <div className="row-actions">
                      <button className="btn btn-ghost btn-sm" onClick={() => handleExtend(lic.id)}>Prolonger</button>
                      <button className="btn btn-ghost btn-sm" onClick={() => handleRevoke(lic.id)}>Révoquer</button>
                      <button className="btn btn-danger btn-sm" onClick={() => handleDelete(lic.id)}>Suppr.</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {totalPages > 1 && (
        <div className="pagination">
          <button className="btn btn-ghost btn-sm" disabled={page <= 1} onClick={() => setPage(page - 1)}>Précédent</button>
          <span className="info">Page {page} / {totalPages} — {total} licence(s)</span>
          <button className="btn btn-ghost btn-sm" disabled={page >= totalPages} onClick={() => setPage(page + 1)}>Suivant</button>
        </div>
      )}

      {showModal && (
        <GenerateModal
          onClose={() => setShowModal(false)}
          onGenerate={handleGenerate}
          generatedKey={generatedKey}
          onCopy={copyKey}
        />
      )}

      {toast && <div className={`toast toast-${toast.type}`}>{toast.msg}</div>}
    </div>
  );
}

function GenerateModal({ onClose, onGenerate, generatedKey, onCopy }: {
  onClose: () => void;
  onGenerate: (form: any) => void;
  generatedKey: string | null;
  onCopy: (key: string) => void;
}) {
  const [form, setForm] = useState({
    customer_name: "",
    customer_email: "",
    edition: "STANDARD",
    expires_at: "",
    max_activations: 1,
    machine_id: "",
  });

  const defaultDate = new Date();
  defaultDate.setFullYear(defaultDate.getFullYear() + 1);
  if (!form.expires_at) form.expires_at = defaultDate.toISOString().substring(0, 10);

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        {generatedKey ? (
          <>
            <h2>Licence générée</h2>
            <p style={{ color: "var(--text-dim)", fontSize: 14, marginBottom: 8 }}>
              Communiquez cette clé à votre client. Elle devra être saisie dans Excel via
              l'onglet Administration {'>'} Licence.
            </p>
            <div className="generated-key">
              <div className="key-text">{generatedKey}</div>
              <button className="copy-btn" onClick={() => onCopy(generatedKey)}>Copier la clé</button>
            </div>
            <div className="modal-actions">
              <button className="btn btn-primary" onClick={onClose}>Terminer</button>
            </div>
          </>
        ) : (
          <>
            <h2>Nouvelle licence</h2>
            <div className="form-group">
              <label>Nom du client</label>
              <input value={form.customer_name} onChange={(e) => setForm({ ...form, customer_name: e.target.value })} placeholder="Entreprise SA" />
            </div>
            <div className="form-group">
              <label>Email du client</label>
              <input type="email" value={form.customer_email} onChange={(e) => setForm({ ...form, customer_email: e.target.value })} placeholder="contact@entreprise.com" />
            </div>
            <div className="form-group">
              <label>Édition</label>
              <select value={form.edition} onChange={(e) => setForm({ ...form, edition: e.target.value })}>
                <option value="STANDARD">Standard</option>
                <option value="PREMIUM">Premium</option>
              </select>
            </div>
            <div className="form-group">
              <label>Date d'expiration</label>
              <input type="date" value={form.expires_at} onChange={(e) => setForm({ ...form, expires_at: e.target.value })} />
            </div>
            <div className="form-group">
              <label>Nombre d'activations max</label>
              <input type="number" min={1} max={100} value={form.max_activations} onChange={(e) => setForm({ ...form, max_activations: parseInt(e.target.value) || 1 })} />
            </div>
            <div className="form-group">
              <label>ID machine (optionnel — laissez vide pour licence flottante)</label>
              <input value={form.machine_id} onChange={(e) => setForm({ ...form, machine_id: e.target.value })} placeholder="ANY" />
            </div>
            <div className="modal-actions">
              <button className="btn btn-ghost" onClick={onClose}>Annuler</button>
              <button className="btn btn-primary" onClick={() => onGenerate(form)}>Générer</button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
