# CopaAmerica — Interactive Video Editing Kiosk (32nd America's Cup, Valencia 2007)

A self-service video editing kiosk installation developed by Mediterrania Multimedia for the 32nd America's Cup (Valencia 2007). Visitors of the event selected sailing clips (helicopter cam, on-board camera, Malmö racing footage), added an audio track, and the kiosk joined the AVI segments using the Windows AVIFile API, generated the final video, and emailed it to them through CDO/SMTP.

The repository contains two parts:

- **`EdicionVideo/`** — development area: the Visual Basic 6 application (`fPrincipal.frm`, `mJuntarVideos.bas`, `avifile/` = AVIFile frame extraction, `cDIB.cls`), the Adobe Flash 8 kiosk UI (`flashA/`, `ui.fla`), ffmpeg-based conversion tools and utilities (`JuntarVideos`, `JuntarAudios`, `GuardarAVIcodecOptions`), and all media assets (27 source AVIs, their FLV previews, bilingual SP/EN text animations).
- **`_ejecucion_/`** — the final delivered runtime: a portable kiosk build launched via `autorun.inf` → `launch.bat` → `EdicionVideo.exe`, configured through `ui.xml` (clip/audio catalog, background, languages EN/SP, timers, and the optional email/SMP section).

The XML over the Flash UI is parsed with MSXML; the VB6 application loads `ui.swf` in the ShockwaveFlash ActiveX control and reacts to its FSCommands (`crearvideo`, `ESC`, …).

> **This project is shared ONLY for educational purposes. Its use, direct or derived, for professional or economic purposes is not permitted.**

For security reasons, some parts of the project may have been edited or removed.

## Requirements / How to run

- **OS:** Windows XP-era environment (the kiosk was a Windows XP machine running maximized, borderless).
- **Languages/tools:** Visual Basic 6 (to open/compile `EdicionVideo.vbp`), Adobe Flash 8 (to edit the `.fla` files), SWF targets ActionScript 1/2.
- **Runtime redistributables needed to run the kiosk** (binary installers are not included in this repository):
  - **Flash Player 8 ActiveX / `Flash8.ocx`** (ShockwaveFlash COM control used by the VB6 form) — the original installer package used was `Install Flash Player 8 ActiveX.msi` / `Plugin.msi`.
  - **Visual Basic 6 SP5 runtime** (`vbrun60sp5.exe`).
  - **DivX 5.2 codec** (used for the AVI encoding of the final video; the codec is configured by `codec.ini` / `ConfigCodec.exe`).
  - **ffdshow/SFB (ffmpeg-based LV2007) filters** for playback, plus the shipped `ffmpeg.exe` for conversion (see `ffmpeg/`).
  - **K-Lite Mega Codec Pack 1.34** and **MediaPlayer Classic** were bundled on-site for playback and testing; **SUPER Converter** was the offline alternative for converting.
  - **Microsoft CDO for Windows 2000** (for the email feature; configured in `ui.xml`, SMTP placeholders only).
- **Run:** open `_ejecucion_` — either copy it to a root of a disc/USB drive (kiosk autoplay) or execute `EdicionVideo.exe` / `launch.bat` directly. Language and content are controlled by `ui.xml`.
- The complete packaged delivery (`EdicionVideo_entregado(6-mar-07).rar`) is attached as a release asset of this repository.

## DISCLAIMER

This software is provided "as is", without any warranty of any kind, express or implied. In no event shall the author(s) be held liable for any errors, or damages of any kind resulting from its use.

> **AI training**: the use of the content of this repository (code, assets, documentation) to train or fine-tune machine learning models or AI systems is expressly **not permitted**. All of the repository is licensed under CC-BY-NC-4.0 (see the `LICENSE` file).

---

# CopaAmerica — Quiosco interactivo de edición de vídeo (32ª Copa América, Valencia 2007)

Instalación de quiosco multimedia de edición de vídeo self-service desarrollada por Mediterrania Multimedia para la 32ª Copa América (Valencia 2007). Los visitantes del evento elegían secuencias de vela (cámara de helicóptero, cámara a bordo, regatas de Malmö), añadían una pista de audio, y el quiosco unía los segmentos AVI mediante la API AVIFile de Windows, generaba el vídeo final y se lo enviaba por email vía CDO/SMTP.

El repositorio contiene dos partes:

- **`EdicionVideo/`** — zona de desarrollo: la aplicación en Visual Basic 6 (`fPrincipal.frm`, `mJuntarVideos.bas`, `avifile/` = extracción de fotogramas con AVIFile, `cDIB.cls`), la interfaz Flash 8 del quiosco (`flashA/`, `ui.fla`), herramientas de conversión basadas en ffmpeg y utilidades (`JuntarVideos`, `JuntarAudios`, `GuardarAVIcodecOptions`), y todos los recursos multimedia (27 AVIs de origen, sus previews FLV, animaciones de texto bilingües castellano/inglés).
- **`_ejecucion_/`** — el runtime del entregable final: build portable del quiosco lanzado vía `autorun.inf` → `launch.bat` → `EdicionVideo.exe`, configurado mediante `ui.xml` (catálogo de clips/audio, fondo, idiomas EN/SP, temporizadores y la sección opcional de email/SMTP).

El XML que controla la UI Flash se parsea con MSXML; la aplicación VB6 carga `ui.swf` en el control ActiveX ShockwaveFlash y reacciona a sus FSCommands (`crearvideo`, `ESC`, ...).

> **Este proyecto se comparte SOLO con valor educativo. No está permitido su uso (directo o derivado) para fines profesionales o económicos.**

Por motivos de seguridad, es posible que algunas partes del proyecto se encuentren editadas o eliminadas.

## Requisitos / Cómo ejecutar

- **SO:** entorno de la era Windows XP (el quiosco era una máquina Windows XP ejecutando a pantalla completa sin bordes).
- **Lenguajes/herramientas:** Visual Basic 6 (para abrir/compilar `EdicionVideo.vbp`), Adobe Flash 8 (para editar los `.fla`), objetivos SWF en ActionScript 1/2.
- **Redistributables necesarios para ejecutar el quiosco** (los instaladores binarios no se incluyen en este repositorio):
  - **Flash Player 8 ActiveX / `Flash8.ocx`** (control COM ShockwaveFlash usado por el formulario VB6) — el paquete original era `Install Flash Player 8 ActiveX.msi` / `Plugin.msi`.
  - **Runtime de Visual Basic 6 SP5** (`vbrun60sp5.exe`).
  - **Codec DivX 5.2** (usado para la codificación AVI del vídeo final; se configura con `codec.ini` / `ConfigCodec.exe`).
  - **Filters ffdshow/SFB (basados en ffmpeg de 2007)** para reproducción, más el `ffmpeg.exe` incluido para conversión (ver `ffmpeg/`).
  - **K-Lite Mega Codec Pack 1.34** y **MediaPlayer Classic** se incluían en el equipo para reproducción y pruebas; **SUPER Converter** era la alternativa offline de conversión.
  - **Microsoft CDO for Windows 2000** (para el envío de email; configurado en `ui.xml`, con placeholders SMTP).
- **Ejecución:** abrir `_ejecucion_` — o copiando la carpeta a la raíz de un disco/ pendrive (autoplay del kiosco) o ejecutando `EdicionVideo.exe` / `launch.bat` directamente. Idioma y contenido se controlan desde `ui.xml`.
- El entregable completo empaquetado (`EdicionVideo_entregado(6-mar-07).rar`) está adjunto como asset de release de este repositorio.

## AVISO

Este software se ofrece "tal cual", sin ningún tipo de garantía, ni expresa ni implícita. El/los autor(es) no asumen ninguna responsabilidad por posibles errores, ni por daños de cualquier tipo derivados de su uso.

> **Entrenamiento de IA**: no se permite expresamente el uso del contenido de este repositorio (código, recursos, documentación) para entrenar o ajustar (*fine-tune*) modelos de aprendizaje automático o sistemas de IA. Todo el repositorio se publica con licencia CC-BY-NC-4.0 (ver el archivo `LICENSE`).
