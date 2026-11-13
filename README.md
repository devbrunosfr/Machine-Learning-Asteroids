# Machine Learning Asteroids

Este relatório apresenta o desenvolvimento de uma solução de Machine Learning 
em ambiente de computação em nuvem para classificação de asteroides próximos 
da Terra como potencialmente perigosos (Potentially Hazardous Asteroids — PHA). 
A solução utiliza dados oficiais da NASA/JPL Small-Body Database (SBDB), 
armazenados em bucket versionado, preparados por meio de uma rotina 
reprodutível e utilizados no treinamento e na avaliação de três algoritmos de 
classificação. O modelo selecionado foi o Gradient Boosting por histogramas, 
registrado como versão champion e disponibilizado por uma API Flask com os 
endpoints /health e /predict. A avaliação final alcançou acurácia de 0,7607, recall de 
0,7412, F1 de 0,2689 e ROC AUC de 0,8415 para a classe PHA. Foram realizados 
testes funcionais, teste de carga com 300 requisições e oito clientes concorrentes, 
monitoramento estruturado, verificação de data drift pelo índice PSI e análise de 
queda de desempenho. O experimento foi executado no modo padrão de bucket 
local emulado, mantendo a mesma interface de armazenamento prevista para uma 
migração ao Google Cloud Storage. O sistema não estima colisão, probabilidade de 
impacto ou dano; classifica o rótulo oficial PHA a partir de elementos orbitais.
