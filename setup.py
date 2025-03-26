from setuptools import setup, find_packages

setup(
    name='dxf_diego_project',
    version='0.1.0',  # Versão inicial
    description='Projeto de visualização, processamento e conversão de arquivos DXF',
    author='Diego', 
    author_email='diego_murari@hotmail.com',
    url='https://github.com/seu_usuario/dxf_diego_project',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    install_requires=[
        'setuptools',
        'wheel',
        'pyparsing>=2.0.1',
        'typing_extensions>=4.6.0',
        'numpy',
        'fonttools',
        'pytest',
        'mypy>=1.11',
        'cython',
        'matplotlib',
        'pyside6',
        'Pillow',
        'PyMuPDF',
        'sphinx',
        'sphinx-rtd-theme',
        'sphinxcontrib-jquery',
        'openpyxl'
    ],
    classifiers=[
        'Development Status :: 3 - Alpha',  # Ou Beta/Production conforme o projeto evolua
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.10',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent'
    ],
    python_requires='>=3.7',
)