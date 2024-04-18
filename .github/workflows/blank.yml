name: Gerar Estudo de Capabilidade

on:
  push:
    branches:
      - main  # Ou a branch que deseja monitorar

jobs:
  gerar_estudo_capabilidade:
    runs-on: ubuntu-latest
    
    steps:
    - name: Checkout do repositório
      uses: actions/checkout@v2
      
    - name: Configurar ambiente Python
      uses: actions/setup-python@v2
      with:
        python-version: '3.x'
      
    - name: Instalar dependências
      run: |
        pip install pandas openpyxl matplotlib
      
    - name: Executar script Python
      run: python gerar_estudo_capabilidade.py  # Substitua pelo nome do seu arquivo Python
