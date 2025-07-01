# Calculator Web

This repository contains a simple HTML calculator for shipping costs and a set of VBA macros for creating an Excel form.

- `index.html` – Web-based shipping cost calculator written in vanilla JavaScript.
- `vba/GenerarFormularioCompleto.bas` – VBA module that programmatically builds a form in Excel and stores submitted values in the active worksheet.

The VBA module can be imported into Excel (Alt+F11) and executed with `GenerarFormularioCompleto` to create the form. The `modFormulario` module created by the macro exposes `AbrirFormulario` to show the generated form.
