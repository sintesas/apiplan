<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Log;

use App\Models\Menu;
use App\Models\Usuario;
use App\Models\UsuarioMenu;
use App\Models\UsuarioRol;

class UsuarioController extends Controller
{
    public function login(Request $request) {
        $p_usuario = $request->get('usuario');
        $p_password = $request->get('password');
        
        try {
            if ($p_password != null) {
                $existe_usuario = Usuario::where('usuario', $p_usuario)->first();

                if ($existe_usuario) {
                    $user = $existe_usuario;

                    if ($user->activo == 'S') {
                        $m_menu = new Menu;
                        $m_usuariomenu = new UsuarioMenu;

                        if (\DB::select('select * from vw_usuarios_roles_privilegios where usuario_id = :id', array('id' => $user->usuario_id)) != null) {
                            $opc = 1;
                        }
                        else {
                            $opc = 2;
                        }

                        $data = array();
                        $data['usuario_id'] = $user->usuario_id;
                        $data['usuario'] = $user->usuario;
                        $data['nombre_completo'] = $user->nombre_completo;
                        $data['email'] = $user->email;
                        $data['menus'] = $m_menu->get_menu_id($m_usuariomenu->getUsuarioMenu($user->usuario_id, $opc));

                        $response = json_encode(array('result' => $data), JSON_NUMERIC_CHECK);
                        $response = json_decode($response);
                        return response()->json(array('user' => $response, 'tipo' => 0));
                    }
                    else {
                        return response()->json(array('mensaje' => "El usuario '" . $user->usuario . "' no se encuentra activo.", 'tipo' => -1, 'codigo' => 2));
                    }
                }
            }
        }
        catch (\PDOException $e) {
            if ($e->getCode() == "08001") {
                return response()->json(array('tipo' => -1, 'mensaje' => "No se puede conectar a la base de datos. Contactar al administrador."));
            }
        }
    }

    public function getUsuarios(Request $request) {
        $model = new Usuario();

        $datos = $model->get_usuarios($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getUsuariosFull(Request $request) {
        $model = new Usuario();

        $datos = $model->get_usuarios_full($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function crearUsuarios(Request $request) {
        $model = new Usuario();

        try {
            $db = $model->crud_usuarios($request, 'C');

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

    public function actualizarUsuarios(Request $request) {
        $model = new Usuario();
        
        try {
            $db = $model->crud_usuarios($request, 'U');

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

    public function getUsuariosRolesById(Request $request) {
        $model = new UsuarioRol();

        $datos = $model->get_usuarios_roles_by_usuario_id($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function crearUsuarioRol(Request $request) {
        $model = new UsuarioRol();

        try {
            $db = $model->crud_usuarios_roles($request, 'C');

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

    public function actualizarUsuarioRol(Request $request) {
        $model = new UsuarioRol();

        try {
            $db = $model->crud_usuarios_roles($request, 'U');

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

    public function getRolPrivilegiosPantalla() {
        $model = new UsuarioRol();

        $datos = $model->get_rol_privilegios_pantalla();

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function getPermisosByUser(Request $request) {
        $model = new Usuario();

        $datos = $model->get_permisos_by_user($request);

        $response = json_encode(array('result' => $datos[0], 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }

    public function eliminarUsuariosRolesId(Request $request) {
        $model = new UsuarioRol();

        try {
            $db = $model->eliminar_usuarios_roles_by_id($request);

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

    public function getRolesByUsuarioId(Request $request) {
        $model = new Usuario();

        $datos = $model->get_roles_by_usuario_id($request);

        $response = json_encode(array('result' => $datos, 'tipo' => 0), JSON_NUMERIC_CHECK);
        $response = json_decode($response);

        return response()->json($response, 200);
    }
}
