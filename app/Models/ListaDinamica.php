<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class ListaDinamica extends Model
{
    use HasFactory;

    protected $table = 'tb_app_nombres_listas';

    protected $primaryKey = 'nombre_lista_id';

    protected $fillable = [
        'nombre_lista,descripcion,nombre_lista_padre_id,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_nombres_listas(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_nombres_listas ?,?,?,?,?,?,?,?", 
                        [
                            $evento,
                            $request->input('nombre_lista_id'),
                            $request->input('nombre_lista'),
                            $request->input('descripcion'),
                            $request->input('nombre_lista_padre_id') == 0 ? null : $request->input('nombre_lista_padre_id'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_nombres_listas(Request $request) {
        $db = DB::select("exec pr_get_app_nombres_listas ?,?",
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 200
                        ]);
        return $db;
    }

    public function get_nombres_listas_full() {
        $db = DB::select("exec pr_get_app_nombres_listas_full");
        return $db;
    }
}
