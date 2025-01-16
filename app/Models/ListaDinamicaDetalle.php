<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class ListaDinamicaDetalle extends Model
{
    use HasFactory;

    protected $table = 'tb_app_listas_dinamicas';

    protected $primaryKey = 'lista_dinamica_id';

    protected $fillable = [
        'nombre_lista_id,lista_dinamica,descripcion,lista_dinamica_padre_id,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_listas_dinamicas(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_listas_dinamicas ?,?,?,?,?,?,?,?,?", 
                        [
                            $evento,
                            $request->input('lista_dinamica_id'),
                            $request->input('nombre_lista_id'),
                            $request->input('lista_dinamica'),
                            $request->input('descripcion'),
                            $request->input('lista_dinamica_padre_id') == 0 ? null : $request->input('lista_dinamica_padre_id'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_listas_dinamicas_by_nombre_lista_id(Request $request) {
        $db = DB::select("exec pr_get_app_listas_dinamicas_by_nombre_lista_id ?", array($request->input('nombre_lista_id')));
        return $db;
    }

    public function get_listas_dinamicas_full() {
        $db = DB::select("exec pr_get_app_listas_dinamicas_full");
        return $db;
    }
}
