
## How to Use

1. Open the Excel file containing the sheets:
   - `ENVIO AWB LOOP`
   - `ENVIO AWB SEM LOOP`
2. Press `Alt + F11` and import `src/Enviar_Email_DHL_Loop.bas`
3. Update the folder path and email addresses in the code
4. Run the macro `Enviar_Email_DHL_Loop`

## Configuration

| Item               | Description                              | Example                        |
|--------------------|------------------------------------------|--------------------------------|
| `CaminhoPasta`     | Folder where PDF/XML files are stored    | `"C:\Temp\HAWB\"`              |
| `EmailPara`        | Main recipients                          | `"destinatario@empresa.com"`   |
| `EmailCC`          | CC recipients                            | `"copia@empresa.com"`          |

## Sheet Structure

**ENVIO AWB LOOP** (main sheet)

| Column | Content              |
|--------|----------------------|
| A      | Reference / ID       |
| B      | AWB number           |
| C      | TE Reference         |
| D      | Extra text for body  |

**ENVIO AWB SEM LOOP** (rows without matching files)
