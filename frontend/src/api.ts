export interface ConvertConfig {
  turnstile_site_key: string;
}

export async function fetchConfig(): Promise<ConvertConfig> {
  const res = await fetch("/api/config");
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error(data.error || "Gagal memuat konfigurasi server.");
  }
  return res.json();
}

export interface ConvertResult {
  blob: Blob;
  filename: string;
  headerCount: number;
  dataCount: number;
}

export async function verifyCaptcha(token: string): Promise<string> {
  const form = new FormData();
  form.append("token", token);
  const res = await fetch("/api/verify", { method: "POST", body: form });
  if (!res.ok) {
    let message = "Verifikasi captcha gagal. Silakan coba lagi.";
    try {
      const data = await res.json();
      if (data.detail) message = data.detail;
    } catch {
      /* ignore */
    }
    throw new Error(message);
  }
  const data = await res.json();
  return data.session as string;
}

export async function convertPdf(
  session: string,
  file: File,
  removeFooter: boolean,
  autoWidth: boolean
): Promise<ConvertResult> {
  const form = new FormData();
  form.append("session", session);
  form.append("file", file);
  form.append("remove_footer", String(removeFooter));
  form.append("auto_width", String(autoWidth));

  const res = await fetch("/api/convert", {
    method: "POST",
    body: form,
  });

  if (!res.ok) {
    let message = "Terjadi kesalahan saat memproses file.";
    try {
      const data = await res.json();
      if (data.detail) message = data.detail;
    } catch {
      /* ignore */
    }
    throw new Error(message);
  }

  const blob = await res.blob();
  const filename =
    res.headers.get("X-Filename") ||
    file.name.replace(/\.pdf$/i, ".xlsx") ||
    "converted.xlsx";
  const headerCount = Number(res.headers.get("X-Header-Count") || "0");
  const dataCount = Number(res.headers.get("X-Data-Count") || "0");

  return { blob, filename, headerCount, dataCount };
}
