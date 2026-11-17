# ✈️ Monitoreo Automatizado de Vuelos BUE -> Barranquilla (Noviembre 2026)

Sistema de monitoreo horario desatendido (24/7 en la nube) para tarifas de vuelos entre Buenos Aires (BUE) y Barranquilla (BAQ), con actualización en vivo del archivo de Google Drive, alertas automáticas ante tarifas menores a **US$ 1.500** y reporte diario a las 20:00 ARG con el Excel adjunto a `urmandiego@gmail.com`.

---

## 🚀 Puesta en Marcha en GitHub Actions (5 minutos)

Este repositorio está preparado para ejecutarse automáticamente cada 1 hora en GitHub Actions de forma 100% gratuita, sin necesidad de tener tu computadora prendida.

### Paso 1: Crear un nuevo repositorio en GitHub
1. Entrá a [GitHub -> New Repository](https://github.com/new).
2. Nombralo (por ejemplo: `vuelos-barranquilla`).
3. Seleccionalo como **Privado**.
4. Crealo y subí esta carpeta:
```bash
git init
git add .
git commit -m "feat: setup flight monitoring with github actions"
git branch -M main
git remote add origin https://github.com/<tu-usuario>/vuelos-barranquilla.git
git push -u origin main
```

---

### Paso 2: Habilitar permisos de escritura en GitHub Actions
Para que GitHub Actions pueda guardar cada actualización en el historial:
1. En tu repositorio, andá a **Settings** -> **Actions** -> **General**.
2. Bajá hasta **Workflow permissions**.
3. Seleccioná **Read and write permissions**.
4. Hacé clic en **Save**.

---

### Paso 3: Configurar los Secrets (Para alertas y correos)
Andá a **Settings** -> **Secrets and variables** -> **Actions** -> **New repository secret**:

| Secret Name | Valor requerido |
| :--- | :--- |
| `SMTP_USER` | Tu cuenta de Gmail: `urmandiego@gmail.com` |
| `SMTP_PASS` | **Contraseña de Aplicación de Google** (16 caracteres). Generala en [Google App Passwords](https://myaccount.google.com/apppasswords) bajo "Seguridad" -> "Verificación en dos pasos" -> "Contraseñas de aplicaciones". |
| `ALERT_EMAIL` | (Opcional, por defecto `urmandiego@gmail.com`). |
| `GDRIVE_SERVICE_ACCOUNT_JSON` | (Opcional) JSON de cuenta de servicio de Google Cloud si deseás que GitHub Actions actualice directamente el ID en Google Drive. Si no se configura, el Excel se actualiza y respalda automáticamente en el repositorio y en los correos adjuntos. |

---

## ⏰ Cronograma y Alertas Automáticas

- **Frecuencia:** Cada 1 hora en punto (`cron: '0 * * * *'`).
- **Alerta Inmediata:** Si cualquiera de las rutas o combinaciones (Lima o Bogotá) baja de **US$ 1.500**, dispara un email de alerta inmediato con el detalle y acceso directo.
- **Resumen Diario:** A las **20:00 hora Argentina**, envía un email con la tabla comparativa del día y el archivo Excel adjunto.
- **Ejecución Manual:** En la pestaña **Actions** de GitHub podés hacer clic en **"Run workflow"** en cualquier momento para forzar un chequeo en vivo.
