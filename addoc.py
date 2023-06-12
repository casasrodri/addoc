import os
import requests
import json
from urllib.parse import quote


class AddocReporteParam:
    def __init__(self) -> None:
        self.data = {
            'ext': 'XLS',
            'SerieDocumental': '0045.30.0.0.0.0.1',
            'Parametros': {
                'SerieDocumental': '0045.30.0.0.0.0.1',
                'ordenarPor': 'ID',
                'ordenAlfabetico': 'DESC',
                'Legajo': '',
                'DNI': '',
                'ApellidoNombre': 'null',
                'COBIS': '',
                'DescripcionProducto': '',
                'NumeroDeCuenta': {'Desde': '', 'Hasta': ''},
                'Observaciones': '',
                'Campo7': '',
                'Campo8': 'null',
                'Campo9': 'null',
                'Campo10': 'null',
                'Campo11': 'null',
                'Campo12': 'null',
                'Campo13': '',
            },
            'documentoBuscar': 'Legajo',
        }
        self.file = {}

    def desde(self, value: int):
        self.data['Parametros']['NumeroDeCuenta']['Desde'] = str(value)
        self.file['desde'] = str(value).replace('/', '')

    def hasta(self, value: int):
        self.data['Parametros']['NumeroDeCuenta']['Hasta'] = str(value)
        self.file['hasta'] = str(value).replace('/', '')

    def dia(self, value: int):
        self.desde(value)
        self.hasta(value)

    def dni(self, value: int):
        self.data['Parametros']['DNI'] = str(value)
        self.file['dni'] = str(value)

    def ic(self, value: int):
        self.data['Parametros']['DescripcionProducto'] = str(value)
        self.file['numic'] = str(value)

    def producto(self, value: int):
        self.data['Parametros']['ApellidoNombre'] = str(value)
        self.file['producto'] = str(value)

    def url(self):
        params = []
        self.data['Parametros'] = quote(
            json.dumps(self.data['Parametros']).replace(' ', '')
        ).replace('/', r'%5C%2F')
        for k, v in self.data.items():
            params.append(f'{k}={v}')

        return '&'.join(params)

    def file_name(self):
        out = []
        for k, v in self.file.items():
            out.append(f'{k}-{v}')
        return '_'.join(out)


class AddocManager:
    session_id = None

    def __init__(self, out_folder: str = '.') -> None:
        requests.packages.urllib3.disable_warnings()
        self.out_folder = out_folder
        self.__mkdir('/')

    def __mkdir(self, relative: str = ''):
        try:
            os.makedirs(self.out_folder + relative)
        except FileExistsError:
            pass

    def __get_session_id(self):
        login = requests.get('https://documentaweb.addoc.com.ar/login.php', verify=False)
        self.session_id = login.headers['Set-Cookie'].split(';')[0]

    def __get(self, url: str):
        return requests.get(
            f'https://documentaweb.addoc.com.ar{url}',
            headers={'Cookie': self.session_id},
            verify=False,
        )

    def login(self, user: str, pwd: str):
        self.__get_session_id()
        validacion = self.__get(f'/ajax/ajax.usuario.php?a=validate&u={user}&p={pwd}')

        estado = json.loads(validacion.text)['estado']

        if estado == 'Error':
            detalle = json.loads(validacion.text)['detalle_estado'].replace('&ntilde;', 'ñ')
            raise Exception(f'Error de autenticación: {detalle}')

        return json.loads(validacion.text)['estado']

    def descargar_excel(self, param: AddocReporteParam):
        xls = self.__get(f'/busqueda.getResultados.php?{param.url()}').content

        self.__mkdir('/Excel')
        path = f"{self.out_folder}/Excel/Reporte_{param.file_name()}.xls"

        with open(path, 'wb') as f:
            f.write(xls)

        return path

    def descargar_legajo(self, id: str):
        if not self.session_id:
            raise Exception(
                'Se deben ingresar las credenciales antes de descargar legajos. Usar el método login(user, pwd).'
            )
        respuesta = self.__get(f'/documento.getImagen.php?imagen={id}&unificado=ok&cliente=0045')

        if respuesta.headers['Content-Type'] == 'application/pdf':
            self.guardar_pdf(id, respuesta.content)
            return 'descargado'
        else:
            return 'no disponible'

    def guardar_pdf(self, id: str, contenido):
        self.__mkdir('/PDF')
        path = f"{self.out_folder}/PDF/{id}.pdf"

        with open(path, 'wb') as f:
            f.write(contenido)
