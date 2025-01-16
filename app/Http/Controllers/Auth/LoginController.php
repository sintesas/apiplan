<?php

namespace App\Http\Controllers\Auth;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;

use Aacotroneo\Saml2\Saml2Auth;
use Tymon\JWTAuth\Facades\JWTAuth;
use App\Models\Menu;
use App\Models\Usuario;
use App\Models\UsuarioMenu;

class LoginController extends Controller
{
    public function saml() {
        $url = "http://localhost:4200";
        // $url = 'https://plancarrera.fac.mil.co';

        if (\Auth::guest()) {
            try {
                $saml2Auth = new Saml2Auth(Saml2Auth::loadOneLoginAuthFromIpdConfig('plan'));
                return $saml2Auth->login(session()->pull('url.intended'));                
            }
            catch (\Exception $e) {
                return abort(503);
            }
        }
        else {
            $user = \Auth::user();
            $user = base64_encode($user->usuario);
            $query = http_build_query([
                'id' => $user,
                'type' => 'granted'
            ]);
            return redirect($url.'/saml/?'. $query);
        }
    }

    public function login(Request $request) {
        $p_usuario = base64_decode($request->get('id'));
        $opc = 0;

        try {
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

                    $usuario = Usuario::find($user->usuario_id);
                    $token = JWTAuth::fromUser($usuario);

                    $response = json_encode(array('result' => $data, 'token' => $token), JSON_NUMERIC_CHECK);
                    $response = json_decode($response);                    

                    return response()->json(array('user' => $response, 'tipo' => 0));
                }
                else {
                    return response()->json(array('mensaje' => "El usuario '" . $user->usuario . "' no se encuentra activo.", 'tipo' => -1, 'codigo' => 2));
                }
            }
        }
        catch (\PDOException $e) {
            if ($e->getCode() == "08001") {
                return response()->json(array('tipo' => -1, 'mensaje' => "No se puede conectar a la base de datos. Contactar al administrador."));
            }
        }
    }

    public function logout() {
        $url = "http://localhost:4200";
        // $url = 'https://plancarrera.fac.mil.co';

        // recover sessionIndex and nameId from session
        $nameId = session('nameId');
        $sessionIndex = session('sessionIndex');

        try {
            $saml2Auth = new Saml2Auth(Saml2Auth::loadOneLoginAuthFromIpdConfig('plan'));
            $saml2Auth->logout($url, $nameId, $sessionIndex);
        }
        catch (\Exception $e) {
            abort(500);
        }
    }
}
