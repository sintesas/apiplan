<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class Perfil extends Model
{
    use HasFactory;

    protected $table = 'tb_app_perfiles';

    protected $primaryKey = 'perfil_id';

    protected $fillable = [
        'nombres,apellidos,correo_electronico,avatar,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_perfiles(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_perfiles ?,?,?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('perfil_id'),
                            $request->input('nombres'),
                            $request->input('apellidos'),
                            $request->input('correo_electronico'),
                            $request->input('avatar'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_perfiles(Request $request) {
        $db = DB::select("exec pr_get_app_perfiles ?,?", 
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 200
                        ]);
        return $db;
    }

    public function get_perfiles_full() {
        $db = DB::select("exec pr_get_app_perfiles_full");
        return $db;
    }
}
