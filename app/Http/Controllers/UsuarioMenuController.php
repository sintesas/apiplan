<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;

use App\Models\UsuarioMenu;

class UsuarioMenuController extends Controller
{
    public function crudAsignarMenus(Request $request) {
        $model = new UsuarioMenu();

        try {
            $db = $model->crud_asignar_menus($request);

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue guardado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }
}
