<!DOCTYPE html>
<html lang="{{ str_replace('_', '-', app()->getLocale()) }}">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Servicio Web Plan de Carrera</title>
    <link rel="icon" href="images/favicon.ico">
    <style>
        body {
            font-family: Helvetica,Arial,sans-serif;
            font-weight: 400;
            font-size: 12px;
        }

        .container {
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            height: 100%;
            margin-top: 100px;
        }

        .logo {
            width: 438px;
            background: #001326;
            padding: 10px;
            text-align: center;
        }

        .logo img {
            width: 80%;
            height: 80%;
        }

        hr {
            margin-top: 20px;
            border: 0;
            border-top: 1px solid #eee;
            height: 0;
            width: 458px;
            box-sizing: content-box;
        }

        .footer {
            font-size: 14px !important;
            text-align: left;
        }
    </style>
</head>
<body>
<div class="container">
    <div class="logo">
        <img src="images/logo_web_fac.png">
    </div>
    <img src="images/fac_logo.png">
    <hr>
    <div class="footer">
        <p>© {{ now()->year }} - Plan de Carrera - Fuerza Aeroespacial Colombiana</p>
    </div>
</div>
</body>
</html>