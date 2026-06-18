# Dashboard dinámico de arriendo FIAS

Estructura para GitHub Pages:

- `Arriendo.html`: dashboard público. No contiene datos hardcodeados.
- `data/control-arriendo.json`: JSON leído por el dashboard.
- `scripts/build-arriendo-data.mjs`: descarga el Excel público desde OneDrive/SharePoint y genera el JSON.
- `.github/workflows/update-control-arriendo.yml`: workflow para actualizar el JSON manualmente o cada 6 horas.
- `package.json`: dependencias y scripts.

El Excel fuente no se sube a GitHub. El script lo descarga desde el enlace público de OneDrive.
