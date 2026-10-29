# Rendimiento de atenciones · Aranda

Aplicativo web (un solo `index.html`) para medir el cumplimiento del SLA, auditar vencidos, dar seguimiento a los planes de acción y generar el informe semanal. Los planes, las justificaciones y los cortes semanales se comparten entre sedes con Supabase; el Excel de Aranda se carga en cada navegador.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | El aplicativo completo |
| `config.js` | URL y clave pública de tu proyecto Supabase (lo editas tú) |
| `supabase/schema.sql` | Crea la tabla y las reglas de seguridad |
| `supabase/migracion_tickets.sql` | Permite guardar los tickets (solo si la base ya existía) |
| `vercel.json` | Cabeceras de seguridad para Vercel |

## 1. Supabase (base de datos y usuarios)

1. Entra a https://supabase.com → **New project**. Elige nombre, región cercana (por ejemplo *South America (São Paulo)*) y una contraseña de base de datos (guárdala).
2. Cuando termine de crearse: **SQL Editor → New query**, pega todo `supabase/schema.sql` y pulsa **Run**. Debe decir *Success*.
3. **Authentication → Users → Add user → Create new user**. Crea un usuario por persona del equipo (correo + contraseña) y marca **Auto Confirm User**.
4. **Authentication → Sign In / Providers → Email**: desactiva **Allow new users to sign up** (así nadie más puede registrarse) y guarda.
5. **Project Settings → API** (o el botón **Connect**): copia el **Project URL** y la clave **anon public**.

> Nunca uses ni publiques la clave `service_role`.

## 2. Configurar la aplicación

Abre `config.js` y reemplaza los dos valores:

```js
window.APP_CONFIG = {
  SUPABASE_URL: "https://xxxxxxxx.supabase.co",
  SUPABASE_ANON_KEY: "eyJ..."
};
```

La clave `anon` es pública por diseño; los datos quedan protegidos por el inicio de sesión y las reglas de `schema.sql`.

## 3. GitHub

1. https://github.com → **New repository** (puede ser privado), por ejemplo `planes-aranda`.
2. **Add file → Upload files** y sube el contenido de esta carpeta (`index.html`, `config.js`, `vercel.json`, `README.md` y la carpeta `supabase`). **Commit changes**.

Con Git en tu computadora:

```bash
git init
git add .
git commit -m "Primera versión"
git branch -M main
git remote add origin https://github.com/TU-USUARIO/planes-aranda.git
git push -u origin main
```

## 4. Vercel

1. https://vercel.com → **Add New… → Project** → importa el repositorio de GitHub.
2. **Framework Preset: Other**. Deja vacíos *Build Command* y *Output Directory*. **Deploy**.
3. Vercel te da una dirección `https://planes-aranda.vercel.app`. Ábrela, ingresa con uno de los usuarios del paso 1 y listo.

## 5. Primer uso

1. Ingresa con tu usuario. Los planes de ejemplo se suben solos a Supabase la primera vez.
2. Carga el Excel de Aranda (botón «Cargar Excel de Aranda»).
3. Desde ese momento, planes, justificaciones y cortes quedan guardados en la nube. Arriba a la derecha verás «Sincronizado hh:mm».

## Actualizar la aplicación

Cambia `index.html` en GitHub (editar o subir de nuevo). Vercel publica la nueva versión automáticamente.

## Cosas a tener en cuenta

- Los **tickets del Excel de Aranda sí se guardan en la nube** (colección `tickets`). Antes de recargar la página espera a ver «Sincronizado hh:mm» arriba a la derecha. Si ves «Error al guardar: …», revisa el mensaje. Para saber qué versión está publicada, abre la consola del navegador (F12) y escribe `APP_BUILD`.
- Si ya tenías la base de datos creada antes de esta versión, ejecuta `supabase/migracion_tickets.sql` una sola vez.
- Si dos personas editan el mismo registro a la vez, gana el último cambio guardado.
- Sin conexión, los cambios quedan en el navegador y se envían al volver el internet.
