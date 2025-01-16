<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Log;

use App\Models\Cargo;
use App\Models\CargoConfiguracion;
use App\Models\CargoGrado;
use App\Models\Area;
use App\Models\Cuerpo;
use App\Models\Educacion;
use App\Models\Especialidad;
use App\Models\CargoExperiencia;
use App\Models\UbicacionCargo;

class CargoController extends Controller
{
    public function getCargos(Request $request) {
        $model = new Cargo();

        $datos = $model->get_cargos($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getCargosFull() {
        $model = new Cargo();

        $datos = $model->get_cargos_full();

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getCargosId(Request $request) {
        $model = new Cargo();

        $datos = $model->get_cargos_by_id($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function crearCargos(Request $request) {
        $model = new Cargo();

        try {
            $db = $model->crud_cargo($request, 'C');

            if ($db) {
                $id = $db[0]->id;

                $response = json_encode(array('mensaje' => 'Fue creado exitosamente.', 'tipo' => 0, 'id' => $id), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function actualizarCargos(Request $request) {
        $model = new Cargo();
        
        try {
            $db = $model->crud_cargo($request, 'U');

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue actualizado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function eliminarCargos(Request $request) {
        $model = new Cargo();

        try {
            $db = $model->eliminar_cargo($request);

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue eliminado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function getDetalleCargos(Request $request) {
        $model = new Cargo();

        $datos = $model->get_detalle_cargos($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getCargosGradosByCargoId(Request $request) {
        $model = new CargoGrado();

        $datos = $model->get_cargos_grados_by_cargo_id($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function crearCargosGrados(Request $request) {
        $model = new CargoGrado();

        try {
            $db = $model->crud_cargos_grados($request, 'C');

            if ($db) {
                $id = $db[0]->id;

                $response = json_encode(array('mensaje' => 'Fue creado exitosamente.', 'tipo' => 0, 'id' => $id), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function actualizarCargosGrados(Request $request) {
        $model = new CargoGrado();

        try {
            $db = $model->crud_cargos_grados($request, 'U');

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue actualizado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function eliminarCargosGradosById(Request $request) {
        $model = new CargoGrado();

        try {
            $db = $model->eliminar_cargos_grados_by_id($request);

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue eliminado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function getCargosConfiguracionByCargoGradoId(Request $request) {
        $model = new CargoConfiguracion();

        $datos = $model->get_cargos_configuracion_by_cargo_grado_id($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getCargosConfiguracionById(Request $request) {
        $model = new CargoConfiguracion();

        $datos = $model->get_cargos_configuracion_by_id($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function crearCargosConfiguracion(Request $request) {
        $model = new CargoConfiguracion();

        try {
            $db = $model->crud_cargos_configuracion($request, 'C');

            if ($db) {
                $id = $db[0]->id;

                $response = json_encode(array('mensaje' => 'Fue creado exitosamente.', 'tipo' => 0, 'id' => $id), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function actualizarCargosConfiguracion(Request $request) {
        $model = new CargoConfiguracion();

        try {
            $db = $model->crud_cargos_configuracion($request, 'U');

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue actualizado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function getHistorialCuerposByCargoGrado(Request $request) {
        $model = new CargoConfiguracion();

        $datos = $model->get_historial_cuerpos_by_cargo_grado($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getHistorialEspecialidadesByCargoGrado(Request $request) {
        $model = new CargoConfiguracion();

        $datos = $model->get_historial_especialidedades_by_cargo_grado($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getHistorialAreasByCargoGrado(Request $request) {
        $model = new CargoConfiguracion();

        $datos = $model->get_historial_areas_by_cargo_grado($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getCargosExperienciasById(Request $request) {
        $model = new CargoExperiencia();

        $datos = $model->get_app_cargos_experiencias_by_id($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function crearCargosExperiencias(Request $request) {
        $model = new CargoExperiencia();

        try {
            $db = $model->crud_cargos_experiencias($request, 'C');

            if ($db) {
                $id = $db[0]->id;

                $response = json_encode(array('mensaje' => 'Fue creado exitosamente.', 'tipo' => 0, 'id' => $id), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function actualizarCargosExperiencias(Request $request) {
        $model = new CargoExperiencia();

        try {
            $db = $model->crud_cargos_experiencias($request, 'U');

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue actualizado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function eliminarCargosExperienciasId(Request $request) {
        $model = new CargoExperiencia();

        try {
            $db = $model->eliminar_cargos_experiencias_by_id($request);

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue eliminado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function getUbicacionCargosId(Request $request) {
        $model = new UbicacionCargo();

        $datos = $model->get_ubicacion_cargos_by_id($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function crearUbicacionCargos(Request $request) {
        $model = new UbicacionCargo();

        try {
            $db = $model->crud_ubicacion_cargos($request, 'C');

            if ($db) {
                $id = $db[0]->id;

                $response = json_encode(array('mensaje' => 'Fue creado exitosamente.', 'tipo' => 0, 'id' => $id), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function actualizarUbicacionCargos(Request $request) {
        $model = new UbicacionCargo();

        try {
            $db = $model->crud_ubicacion_cargos($request, 'U');

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue actualizado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function eliminarUbicacionCargosId(Request $request) {
        $model = new UbicacionCargo();

        try {
            $db = $model->eliminar_ubicacion_cargos_by_id($request);

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue eliminado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function getListasNiveles() {
        $model = new Cargo();

        $datos = $model->get_listas_niveles();

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function consultarRutasPorCargo($cargo_id) {
        $rutasAsociadas = DB::table('tb_app_rutas')
            ->where('cargo_id', $cargo_id)
            ->pluck('ruta_carrera_id');
    
        return response()->json($rutasAsociadas, 200);
    }

    public function consultarCargosPorCargo($cargo_previo_id) {
        $rutasAsociadas = DB::table('vw_app_cargos_experiencias')
            ->where('cargo_previo_id', $cargo_previo_id)
            ->pluck('cargo_id', 'grado');

        $resultadoOrdenado = [];
        foreach ($rutasAsociadas as $grado => $cargo_id) {
            $resultadoOrdenado[$cargo_id] = $grado;
        }
        
        return response()->json($resultadoOrdenado, 200);
    }

    public function consultarUbicacionesPorCargo($cargo_id) {
        $ubicacionAsociadas = DB::table("vw_app_ubicacion_cargos")
            ->where('cargo_jefe_inmediato_id', $cargo_id)
            ->pluck('cargo_id', 'grado');

        $resultadoOrdenado = [];
        foreach ($ubicacionAsociadas as $grado => $cargo_id) {
            $resultadoOrdenado[$cargo_id] = $grado;
        }

        return response()->json($resultadoOrdenado, 200);
    }
}
