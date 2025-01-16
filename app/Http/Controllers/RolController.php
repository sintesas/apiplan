<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;

use App\Models\Rol;
use App\Models\RolPrivilegio;

class RolController extends Controller
{
    public function getRoles(Request $request) {
        $model = new Rol();

        $datos = $model->get_roles($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getRolesActivos() {
        $model = new Rol();

        $datos = $model->get_roles_activos();

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function crearRoles(Request $request) {
        $model = new Rol();

        try {
            $db = $model->crud_roles($request, 'C');

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

    public function actualizarRoles(Request $request) {
        $model = new Rol();

        try {
            $db = $model->crud_roles($request, 'U');

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue creado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function getRolPrivilegiosById(Request $request) {
        $model = new RolPrivilegio();

        $datos = $model->get_roles_privilegios_by_rol_id($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function crearRolPrivilegios(Request $request) {
        $model = new RolPrivilegio();

        try {
            $db = $model->crud_roles_privilegios($request, 'C');

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

    public function actualizarRolPrivilegios(Request $request) {
        $model = new RolPrivilegio();

        try {
            $db = $model->crud_roles_privilegios($request, 'U');

            if ($db) {
                $response = json_encode(array('mensaje' => 'Fue creado exitosamente.', 'tipo' => 0), JSON_NUMERIC_CHECK);
                $response = json_decode($response);

                return response()->json($response, 200);
            }
        }
        catch (Exception $e) {
            return response()->json(array('tipo' => -1, 'mensaje' => $e));
        }
    }

    public function eliminarRolPrivilegiosById(Request $request) {
        $model = new RolPrivilegio();

        try {
            $db = $model->eliminar_roles_privilegios_by_id($request);

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

    public function getModulos() {
        $model = new RolPrivilegio();

        $datos = $model->get_modulos();

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }
}
