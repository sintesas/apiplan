<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Log;

use App\Models\ListaDinamica;
use App\Models\ListaDinamicaDetalle;

class ListaDinamicaController extends Controller
{
    public function getNombresListas(Request $request) {
        $model = new ListaDinamica();

        $datos = $model->get_nombres_listas($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getNombresListasFull() {
        $model = new ListaDinamica();

        $datos = $model->get_nombres_listas_full();

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function crearNombresListas(Request $request) {
        Log::info($request);

        $model = new ListaDinamica();

        try {
            $db = $model->crud_nombres_listas($request, 'C');

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

    public function actualizarNombresListas(Request $request) {
        Log::info($request);

        $model = new ListaDinamica();

        try {
            $db = $model->crud_nombres_listas($request, 'U');

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

    public function getListasDinamicasFull() {
        $model = new ListaDinamicaDetalle();

        $datos = $model->get_listas_dinamicas_full();

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getListasDinamicas(Request $request) {
        $model = new ListaDinamicaDetalle();

        $datos = $model->get_listas_dinamicas_by_nombre_lista_id($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function crearListasDinamicas(Request $request) {
        Log::info($request);

        $model = new ListaDinamicaDetalle();

        try {
            $db = $model->crud_listas_dinamicas($request, 'C');

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

    public function actualizarListasDinamicas(Request $request) {
        Log::info($request);
        
        $model = new ListaDinamicaDetalle();

        try {
            $db = $model->crud_listas_dinamicas($request, 'U');

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
}
