import { useEffect, useRef, useState } from "react";
import { Turnstile } from "@marsidev/react-turnstile";
import { convertPdf, fetchConfig, verifyCaptcha } from "./api";

export default function App() {
  const [siteKey, setSiteKey] = useState<string | null>(null);
  const [configError, setConfigError] = useState<string | null>(null);
  const [session, setSession] = useState<string | null>(null);
  const [verifying, setVerifying] = useState(false);

  const [file, setFile] = useState<File | null>(null);
  const [removeFooter, setRemoveFooter] = useState(true);
  const [autoWidth, setAutoWidth] = useState(true);

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<{
    filename: string;
    headerCount: number;
    dataCount: number;
    url: string;
  } | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const [tsKey, setTsKey] = useState(0);

  useEffect(() => {
    fetchConfig()
      .then((cfg) => setSiteKey(cfg.turnstile_site_key))
      .catch((err) => setConfigError(err.message));
  }, []);

  useEffect(() => {
    return () => {
      if (result?.url) URL.revokeObjectURL(result.url);
    };
  }, [result]);

  async function handleToken(token: string) {
    setVerifying(true);
    setError(null);
    try {
      const s = await verifyCaptcha(token);
      setSession(s);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Verifikasi captcha gagal.");
      setSession(null);
      setTsKey((k) => k + 1);
    } finally {
      setVerifying(false);
    }
  }

  async function handleConvert() {
    if (!session || !file) return;
    setLoading(true);
    setError(null);
    try {
      const res = await convertPdf(session, file, removeFooter, autoWidth);
      const url = URL.createObjectURL(res.blob);
      setResult({
        filename: res.filename,
        headerCount: res.headerCount,
        dataCount: res.dataCount,
        url,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Terjadi kesalahan.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="page">
      <div className="card">
        <h1>PDF ke Excel Converter - SIPD-RI</h1>

        {configError && <div className="alert alert-error">{configError}</div>}

        {!session && !configError && (
          <div className="turnstile-wrap">
            <p className="caption">Silakan verifikasi captcha terlebih dahulu</p>
            {siteKey ? (
              <Turnstile
                key={tsKey}
                siteKey={siteKey}
                onSuccess={handleToken}
                onError={() => {
                  setSession(null);
                  setError("Verifikasi captcha gagal. Silakan coba lagi.");
                }}
                onExpire={() => {
                  setSession(null);
                  setError("Captcha kedaluwarsa. Silakan verifikasi kembali.");
                }}
              />
            ) : (
              <p className="caption">Memuat captcha…</p>
            )}
            {verifying && (
              <p className="caption">Memverifikasi captcha…</p>
            )}
          </div>
        )}

        {session && (
          <>
            <div className="field">
              <label className="file-label" htmlFor="pdf-input">
                Pilih File PDF
              </label>
              <input
                id="pdf-input"
                ref={fileInputRef}
                type="file"
                accept="application/pdf,.pdf"
                onChange={(e) => {
                  setFile(e.target.files?.[0] ?? null);
                  setResult(null);
                }}
              />
            </div>

            <div className="checks">
              <label className="check">
                <input
                  type="checkbox"
                  checked={removeFooter}
                  onChange={(e) => setRemoveFooter(e.target.checked)}
                />
                Hapus footer &quot;SIPD-RI***&quot;
              </label>
              <label className="check">
                <input
                  type="checkbox"
                  checked={autoWidth}
                  onChange={(e) => setAutoWidth(e.target.checked)}
                />
                Auto lebar kolom Excel
              </label>
            </div>

            {file && (
              <div className="alert alert-info">File terpilih: {file.name}</div>
            )}

            {!file && (
              <div className="alert alert-warning">
                Silakan upload file PDF terlebih dahulu.
              </div>
            )}

            <button
              className="btn btn-primary"
              disabled={!file || loading}
              onClick={handleConvert}
            >
              {loading ? "Sedang memproses…" : "Konversi ke Excel"}
            </button>

            {error && <div className="alert alert-error">{error}</div>}

            {result && (
              <div className="result">
                <div className="alert alert-success">
                  Konversi berhasil! Header: {result.headerCount} baris, Data:{" "}
                  {result.dataCount} baris
                </div>
                <a
                  className="btn btn-download"
                  href={result.url}
                  download={result.filename}
                >
                  Download File Excel
                </a>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
