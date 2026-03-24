# Analise de Consistencia - Labor Rural

Software desktop para analise automatica de inconsistencias em planilhas MIMC.

## Download

Acesse a aba **Releases** e baixe `Analise_Consistencia_LaborRural.exe`.  
Execute diretamente no Windows. Nenhuma instalacao necessaria.

## Verificacoes realizadas

| Aba | Verificacoes |
|-----|-------------|
| INVENTARIO | Valores fora de R$100-R$500k, data fabricacao posterior a aquisicao |
| PRODUCAO | Rateio com talhoes faltando, valores divergentes |
| DESPESAS | Atividades sem M.O., manutencao como Administracao, R$/ha acima de 5k, valor unitario fora do padrao, rateio incompleto, lancamentos identicos duplicados, recorrencia administrativa |
| VENDAS | Preco de venda acima de R$100/sc |

## Publicar nova versao

```bash
git add .
git commit -m "descricao das mudancas"
git push origin main
git tag v1.x.x
git push origin v1.x.x
```

Acesse GitHub > aba **Actions** e aguarde ~8 minutos. O `.exe` aparece em **Releases**.
