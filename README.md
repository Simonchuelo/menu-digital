# 🍔 Menú Digital (GitHub Pages + Google Sheets)

Sistema de menú digital con pedidos por WhatsApp, **100% alojado en GitHub Pages**
(sin servidor propio). Los datos se guardan en una **Google Sheet** gratuita.

## Páginas
- **Cliente** → `cliente/index.html` — el cliente elige producto/cantidad y hace el pedido
- **Admin** → `admin/index.html` — panel de cocina: pedidos pendientes/entregados, confirmar pago, gestionar menú, resumen de ventas

## Características
- ✅ Cliente: fotos de producto, precio, cantidad, carrito, checkout
- ✅ Pago: sube captura (vía imgbb) o marca "ya pagué" / efectivo
- ✅ Admin: alerta con sonido de pedido nuevo, botón WhatsApp con pedido escrito, confirmar pago, estados (pendiente → en preparación → entregado)
- ✅ Gestionar menú: agregar/editar/ocultar/borrar productos con foto
- ✅ Resumen de ventas del día (por producto y método de pago)

## Archivos
| Archivo | Qué es |
|---------|--------|
| `Code.gs` | Backend. Se pega en Google Apps Script (ver INSTRUCCIONES) |
| `config.js` | Config: `API_URL` (Apps Script), `WPP_ADMIN`, `IMGBB_KEY` |
| `cliente/index.html` | Menú del cliente |
| `admin/index.html` | Panel de administración |
| `INSTRUCCIONES.txt` | Guía paso a paso de configuración |

## Configuración rápida
Leé **`INSTRUCCIONES.txt`**. En resumen:
1. Creá una Google Sheet → Apps Script → pegá `Code.gs` → desplegá como "Web app" → copiá la URL
2. Creá una clave gratuita en https://api.imgbb.com
3. Poné ambas cosas en `config.js`
4. Subí esta carpeta a un repo GitHub y activá **GitHub Pages** (branch: main, root)

El QR que hagas debe apuntar a la URL del cliente: `https://TUUSUARIO.github.io/menu-digital/cliente/`
