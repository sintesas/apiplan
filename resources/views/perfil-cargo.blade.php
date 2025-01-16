<!DOCTYPE html>
<html lang="{{ str_replace('_', '-', app()->getLocale()) }}">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{{ $data['cargo'] }}</title>
    <link rel="icon" href="images/favicon.ico">
    <style type="text/css">
        @page {
		    margin: 0cm;
            size: letter portrait;
	    }
        body {
            font-family: "Helvetica Neue", Helvetica;;
            font-size: 12px;
            margin: 4.5cm 2cm 2cm 2cm;
        }
        img {
            height: 75px;
        }
        header, footer {
            position: fixed;
            left: 0cm;
            right: 0cm;
            width: 100%;
        }
        header {
            top: 0cm;            
            height: 3cm;
        }
        footer {
            bottom: 0cm;
            height: 2cm;
        }
        .header-content, .footer-content {
            margin-left: 2cm;
            margin-right: 2cm;
        }
        .header-content {
            margin-top: 1cm;
        }
        .main {
            width: 100%;
            font-size: 13px;
        }
        .table1, .table2, .table3 {
            width: 100%;
            border-collapse: collapse;
        }
        .table1 td, .table2 th, .table2 td, .table3 th, .table3 td {
            border: 1px solid black;
        }
        .table1 .img2 {
            width: 120px;
            height: 80px;
        }
        .table2 th {
            font-size: 9px;
        }
        .table2 td {
            font-size: 8px;
            padding: 2px;
            text-align: center;
        }
        .table2 th, .table3 th {
            background: lightgray;
        }
        .table3 td {
            padding: 5px;
        }
        .center {
            text-align: center;
        }
    </style>
</head>
<body>
    <header>
        <div class="header-content">
            <table class="table1">
                <tr>
                    <td style="text-align: center; width: 150px">
                        <img class="img1" src="../images/logo1.png">
                    </td>
                    <td style="text-align: center; width: auto">
                        <p><strong>FUERZA AEROESPACIAL COLOMBIANA</strong></p>
                        <hr>
                        <p><strong>PLAN CARRERA FAC</strong></p>
                        <p><strong>PERFIL DEL CARGO</strong></p>
                    </td>
                    <td style="text-align: center; width: 150px">
                        <img class="img2" src="../images/logo2.png">
                    </td>
                </tr>
            </table>
        </div>
    </header>
    <footer>
        <div class="footer-content"></div>
    </footer>
    <div class="main">
        <p><strong>DENOMINACIÓN DEL CARGO:</strong>&nbsp;{{ $data['cargo'] }}</p>
        <p><strong>GRADO:</strong>&nbsp;{{ $data['grado'] }}</p>
        <div class="center"><strong>UBICACIÓN DEL CARGO / PUESTOS DE TRABAJO / JEFE INMEDIATO</div><br />
        <table class="table2">
            <thead>
                <tr>
                    <th>Nivel 1</th>
                    <th>Nivel 2</th>
                    <th>Nivel 3</th>
                    <th>Nivel 4</th>
                    <th>Nivel 6</th>
                    <th>Nivel 7</th>
                    <th>Nivel 8</th>
                    <th>Puesto Cantidad</th>
                    <th>Cargo Jefe Inmediato</th>
                </tr>
            </thead>
            <tbody>
                @forelse($data['ubicacion'] as $item)
                <tr>
                    <td>{{ $item->nivel1 ?? 'N/A' }}</td>
                    <td>{{ $item->nivel2 ?? 'N/A' }}</td>
                    <td>{{ $item->nivel3 ?? 'N/A' }}</td>
                    <td>{{ $item->nivel4 ?? 'N/A' }}</td>
                    <td>{{ $item->nivel5 ?? 'N/A' }}</td>
                    <td>{{ $item->nivel6 ?? 'N/A' }}</td>
                    <td>{{ $item->nivel7 ?? 'N/A' }}</td>
                    <td>{{ $item->puesto_cantidad }}</td>
                    <td>{{ $item->cargo_jefe_inmediato ?? 'N/A' }}</td>
                </tr>
                @empty
                <tr>
                    <td colspan="9" style="text-align: center">No hay registros</td>
                </tr>
                @endforelse
            </tbody>
        </table>
        <br />
        <div class="center"><strong><strong>CUERPOS - ESPECIALIDADES - ÁREAS DE CONOCIMIENTO</div><br />
        <table class="table3">
            <thead>
                <tr>
                    <th>CUERPOS</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td style="text-align: justify !important">{{ $data['cuerpo'] ?? 'No hay información' }}</td>
                </tr>
            </tbody>            
        </table>
        <table class="table3">
            <thead>
                <tr>
                    <th>ESPECIALIDADES</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td style="text-align: justify !important">{{ $data['especialidad'] ?? 'No hay información' }}</td>
                </tr>
            </tbody>            
        </table>
        <table class="table3">
            <thead>
                <tr>
                    <th>ÁREAS DE CONOCIMIENTO</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td style="text-align: justify !important">{{ $data['area'] ?? 'No hay información' }}</td>
                </tr>
            </tbody>            
        </table>
        <br />
        <div class="center"><strong><strong>EDUCACIÓN FORMAL</div><br />
        <table class="table3">
            <tbody>
                <tr>
                    <td style="text-align: justify !important">{{ $data['educacion'] ?? 'No hay información' }}</td>
                </tr>
            </tbody>            
        </table>
        <br />
        <div class="center"><strong><strong>CONOCIMIENTOS ADICIONALES CERTIFICADOS</div><br />
        <table class="table3">            
            <tbody>
                <tr>
                    <td style="text-align: justify !important">{{ $data['conocimiento'] ?? 'No hay información' }}</td>
                </tr>
            </tbody>            
        </table>
        <br />
        <div class="center"><strong><strong>EXPERIENCIA DE PROCESO</div><br />
        <table class="table3">            
            <tbody>
                <tr>
                    <td style="text-align: justify !important">{{ $data['experiencia'] ?? 'No hay información' }}</td>
                </tr>
            </tbody>            
        </table>
        <br />
        <div class="center"><strong><strong>COMPETENCIAS</div><br />
        <table class="table3">
            <tbody>
                <tr>
                    <td style="text-align: justify !important">{{ $data['competencia'] ?? 'No hay información' }}</td>
                </tr>
            </tbody>            
        </table>
        <br />
        <div class="center"><strong>EXPERIENCIA: CARGOS PREVIOS A DESEMPEÑAR</div><br />
        <table class="table3">
            <thead>
                <tr>
                    <th style="text-align: center">Cargos</th>
                    <th style="text-align: center">Número de años</th>
                    <th style="text-align: center">Número de meses</th>
                </tr>
            </thead>
            <tbody>
                @forelse($data['expCar'] as $item)
                <tr>
                    <td style="text-align: center">{{ $item->cargo_previo }}</td>
                    <td style="text-align: center">{{ $item->anio }}</td>
                    <td style="text-align: center">{{ $item->mes }}</td>
                </tr>
                @empty
                <tr>
                    <td colspan="3" style="text-align: center">No hay registros</td>
                </tr>
                @endforelse
            </tbody>
        </table>
        <br />
        <div class="center"><strong><strong>PROPÓSITO Y FUNCIONES DEL CARGO</div><br />
        <table class="table3">
            <tbody>
                <tr>
                    <td style="text-align: justify !important"><?php echo $data['requisito_cargo']; ?></td>
                </tr>
            </tbody>            
        </table>
        <br />
    </div>    
</body>
</html>