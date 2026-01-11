# 🏨 recepcion2026‑reservas  
**Módulo independiente de reservas — parte del ecosistema recepcion2026**

Este repositorio contiene el **sistema de reservas** del proyecto administrativo hotelero *recepcion2026*.  
Nació como una separación lógica del repositorio original, donde convivían reservas, consumos y automatizaciones.  
Hoy funciona como módulo autónomo, limpio y preparado para evolucionar hacia una base de datos real.

---

## 🎯 Objetivo del módulo
Gestionar **reservas, disponibilidad y estadías** de manera simple, clara y extensible.

Incluye:

- Procesamiento de reservas  
- Limpieza y normalización de grillas  
- Manejo de archivos históricos  
- Scripts auxiliares para automatizar tareas  
- Preparación para futura migración a SQLite  

---

## 📁 Estructura actual del repositorio

recepcion2026-reservas/
│
├── Grilla de Pax 2030.xlsx
├── GRILLA_DE_PAX_RESPALDO_HISTORICO.ods
├── limpiar_grillas_pisos.py
├── procesar_reservas.py
├── procesar_reservas_old.py
└── README.md

> Esta estructura irá evolucionando hacia un formato modular con carpetas `core/`, `data/`, `templates/` y `tests/`.

---

## 🧠 Contexto histórico
Este módulo contiene **los archivos más antiguos del sistema**, creados antes del desarrollo del módulo de consumos.  
Por eso se separó aquí todo lo relacionado con reservas, mientras que lo más reciente vive en:

- `recepcion2026-consumos` → módulo de consumos  
- `recepcion2026` → automatizaciones, estadísticas y orquestación general  

---

## 🚀 Roadmap

### Próximos pasos
- Crear estructura modular (`core/`, `data/`, `templates/`)  
- Migrar CSV a **SQLite**  
- Implementar capa de acceso a datos  
- Agregar tests unitarios  
- Documentar flujos de trabajo  
- Integrar este módulo con el repo principal `recepcion2026`  

### Futuro
- Dashboard de disponibilidad  
- API interna para comunicación entre módulos  
- Interfaz web ligera para reservas  

---

## 🛠️ Requisitos
- Python 3.10+  
- Librerías estándar (sin dependencias externas por ahora)  
- Archivos CSV de reservas y grillas  

---

## 📦 Instalación y uso

Clonar el repositorio:

git clone https://github.com/xpablodaniel/recepcion2026-reservas

´´´python

	cd recepcion2026-reservas

Ejecutar el procesador de reservas:
´´´python

	python3 procesar_reservas.py


Ejecutar limpieza de grillas:
´´´python

	python3 limpiar_grillas_pisos.py


---

## 🤝 Contribuciones
Este proyecto está en evolución activa.  
Toda mejora, issue o sugerencia es bienvenida.

---

## 🧑‍💻 Autor
Proyecto desarrollado por **Pablo Daniel**, como parte del ecosistema administrativo hotelero *recepcion2026*.
