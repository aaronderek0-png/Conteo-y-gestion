import os
import tempfile
_T = tempfile.mkdtemp()
os.environ.setdefault('MISHI_ESTADO_DIR', _T)
os.environ.setdefault('MISHI_LOG_DIR', os.environ['MISHI_ESTADO_DIR'])
from pathlib import Path
import sys, unittest, tempfile, json
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import invoice_parser as ip
import telegram_ingestor as tg
import procesar_lote_wpp as lote
import imputacion_facturas as imp

class Clasificacion(unittest.TestCase):
    def test_remito_encabezado(self):
        self.assertEqual(ip.tipo_comprobante('REMITO N° 0001-123\nProveedor\nTotal 100'), 'Remito')
    def test_factura_con_referencia(self):
        self.assertEqual(ip.tipo_comprobante('FACTURA A\nREMITO: 0001/00000001'), 'Factura')
    def test_factura_c(self):
        self.assertEqual(ip.tipo_comprobante('FACTURA C\nComp. Nro: 57'), 'Factura')
    def test_factura_c_letra_separada(self):
        self.assertEqual(ip.tipo_comprobante('ORIGINAL\nC\nFACTURA\nSORRENTINO\nCOD. 011'), 'Factura')
    def test_factura_raices(self):
        self.assertEqual(ip.tipo_comprobante('A\nCODIGO: 001\nFACTURA -0001/00062530\nREMITO: 0001/00009071'), 'Factura')
    def test_no_valido(self):
        self.assertEqual(ip.tipo_comprobante('DOCUMENTO NO VALIDO COMO FACTURA'), 'Remito')
    def test_desconocido(self):
        self.assertEqual(ip.tipo_comprobante('PRESUPUESTO\nTotal 100'), 'Desconocido')
    def test_nc(self):
        self.assertEqual(ip.tipo_comprobante('NOTA DE CRÉDITO A\nFactura asociada 123'), 'NC')
    def test_remito_sin_local(self):
        self.assertEqual(lote.estado_de({'tipo':'Remito'}, {'local':None}), 'REMITO_SIN_LOCAL_REVISAR')
    def test_remito_sin_precio(self):
        c={'items':[{'descripcion_original':'Apio','cantidad':2,'precio_unitario':None}], 'numero_factura':'1'}
        self.assertEqual(imp.procesar_remito(c)['total_impreso'],0)
    def test_peso_recibido_y_anulado(self):
        c={'items':[{'descripcion_original':'Apio','cantidad':2,'precio_unitario':100,'unidad_detectada':'KG'}], 'numero_factura':'1', 'anotaciones_manuscritas':['APIO 1,090 KG']}
        r=imp.procesar_remito(c); r['local']='Congreso'
        rows=imp.filas_remitos_pendientes([r])
        self.assertEqual(rows[0][6],2)
        self.assertEqual(rows[0][7],1.09)
        r['renglones'][0]['anulado']=True
        self.assertEqual(imp.filas_remitos_pendientes([r])[0][7],0)
    def test_remitos_sin_numero_no_colisionan(self):
        c={'items':[{'descripcion_original':'Apio','cantidad':2,'precio_unitario':100}], 'proveedor_detectado':'Proveedor'}
        a=imp.procesar_remito({**c,'archivo':'a.jpg'}); b=imp.procesar_remito({**c,'archivo':'b.jpg'})
        self.assertNotEqual(imp.filas_remitos_pendientes([a])[0][0],imp.filas_remitos_pendientes([b])[0][0])
    def test_destino_por_renglon(self):
        comp={'archivo':'wigen.jpg','tipo':'Remito','local_destino_base':'Congreso','items':[
            {'cantidad':10,'local_destino':'Congreso'}, {'cantidad':2,'local_destino':'Congreso'},
            {'cantidad':1,'local_destino':'Lafinur'}]}
        parts=lote.separar_destinos_remito(comp)
        self.assertEqual(len(parts),2)
        self.assertEqual(parts[0]['local_destino_renglon'],'Congreso')
        self.assertEqual([x['cantidad'] for x in parts[0]['items']],[10,2])
        self.assertEqual(parts[1]['local_destino_renglon'],'Lafinur')
        self.assertEqual([x['cantidad'] for x in parts[1]['items']],[1])
    def test_destino_parcial_sin_base_revisar(self):
        comp={'archivo':'x.jpg','tipo':'Remito','items':[{'cantidad':10},{'cantidad':1,'local_destino':'Lafinur'}]}
        self.assertIsNone(lote.separar_destinos_remito(comp)[0]['local_destino_renglon'])
    def test_deshacer_no_borra_registro_de_proveedores(self):
        with patch.object(imp,'_gspread') as api:
            self.assertFalse(imp.deshacer_push({'escrito':True,'tipo':'remitos_proveedor','detalle':[]})['deshecho'])
            api.assert_not_called()

class API:
    def __init__(self): self.enviados=[]; self.respuestas=[]
    def get_updates(self, offset, timeout=0):
        return [{'update_id':10,'message':{'chat':{'id':1},'message_id':7,'document':{'file_id':'a','file_name':'factura.jpg','mime_type':'image/jpeg'}}}]
    def descargar(self, fid, dest): dest.write_bytes(b'contenido'); return dest
    def responder(self, chat, text, mid=None): self.respuestas.append(text)

class Telegram(unittest.TestCase):
    def preparar(self, tmp):
        return [patch.object(tg,'BASE_DIR',tmp),patch.object(tg,'ENTRADA',tmp/'entrada'),patch.object(tg,'STATE_JSON',tmp/'state.json')]
    def test_dry_run_no_consumo_no_push(self):
        with tempfile.TemporaryDirectory() as d:
            tmp=Path(d); state=tmp/'state.json'; state.write_text(json.dumps({'offset':1,'sha256':{}}))
            before=state.read_bytes(); api=API()
            patches=self.preparar(tmp)
            with patches[0],patches[1],patches[2],patch.object(lote,'ejecutar',return_value=({'push_drive':{}},None,None,None)) as run,patch.object(tg,'texto_respuesta',return_value='prueba'):
                tg.pasada(api,escribir=False,push=True)
                self.assertFalse(run.call_args.args[2].push_drive)
                self.assertEqual(state.read_bytes(),before)
                self.assertFalse(api.respuestas)
    def test_error_conserva_offset(self):
        with tempfile.TemporaryDirectory() as d:
            tmp=Path(d); state=tmp/'state.json'; state.write_text(json.dumps({'offset':1,'sha256':{}})); api=API()
            patches=self.preparar(tmp)
            with patches[0],patches[1],patches[2],patch.object(lote,'ejecutar',side_effect=RuntimeError('fallo simulado')):
                tg.pasada(api)
                result=json.loads(state.read_text())
                self.assertEqual(result['offset'],1)
                self.assertEqual(result['sha256'],{})
    def test_success_commits(self):
        with tempfile.TemporaryDirectory() as d:
            tmp=Path(d); state=tmp/'state.json'; state.write_text(json.dumps({'offset':1,'sha256':{}})); api=API()
            patches=self.preparar(tmp)
            with patches[0],patches[1],patches[2],patch.object(lote,'ejecutar',return_value=({'push_drive':{}},None,None,None)),patch.object(tg,'texto_respuesta',return_value='ok'):
                tg.pasada(api)
                result=json.loads(state.read_text())
                self.assertEqual(result['offset'],11)
                self.assertEqual(len(result['sha256']),1)

if __name__=='__main__': unittest.main(verbosity=2)
