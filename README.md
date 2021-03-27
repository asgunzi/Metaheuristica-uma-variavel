# Metaheuristica-uma-variavel
Ex de metaheuristica simples, de uma variável, em VBA



A ideia aqui é escrever uma série de metaheurísticas de dificuldade crescente.

Começando do caso mais simples possível. Variável unidimensional “x”.

Defino um range de valores, no caso, x entre 0 e 100.


varmin = 0 'Range Mínimo
varmax = 100 'Range Máximo


Seja a função objetivo –x^3 + 20*x^2 + 100, mostrado no gráfico abaixo.

Private Function FO(ByVal var) As Double
    'Implementa a função objetivo
    FO = -var ^ 3 + 20 * var ^ 2 + 100

End Function

 
![](https://ferramentasexcelvba.files.wordpress.com/2021/03/curva.png)

Este algoritmo vai: 
 - escolher um valor aleatório entre 0 e 100
- ponderar o valor aleatório com a melhor solução até agora
- lambda controla o mix valor aleatório x melhor solução até agora. No início lambda é pequeno, depois vai aumentando
- função objetivo que avalia a solução
- salva o histórico



For iter = 1 To Niteracoes 'Numero de iteracoes
    lambda = iter / Niteracoes
        
    'Sorteia um valor
    sorteia var, varmin, varmax, varbest, lambda
    
    'Avalia a função atual
    FOatual = FO(var)
    
    'Salva a melhor (maximizando)
    If FOatual > FOBest Then
        FOBest = FOatual
        varbest = var
    End If

    FOhist(iter, 1) = FOBest

Next iter



Resultado. Em apenas 100 iterações, já chegou num valor excelente: x = 13,3

Você pode mudar a FO, ranges, etc.

Próximo passo é fazer um exemplo multidimensional, etc.


