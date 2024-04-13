from .xml2qgis import XML2QGISPlugin
def classFactory(iface):
    return XML2QGISPlugin(iface)