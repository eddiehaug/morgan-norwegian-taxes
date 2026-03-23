from importlib.metadata import version, PackageNotFoundError

try:
    __version__ = version("espp2")
except PackageNotFoundError:
    __version__ = "2025.1"
