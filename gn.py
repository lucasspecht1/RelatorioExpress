def buscar(palavra):
    
    import requests
    from bs4 import BeautifulSoup
    
    response = requests.get(f'https://news.google.com/rss/search?q={palavra}&hl=pt-BR')

    content = response.content

    site = BeautifulSoup(content, 'html.parser')

    meses_a = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    meses = ['/01/','/02/','/03/','/04/','/05/','/06/','/07/','/08/','/09/','/10/','/11/','/12/']
    #print(site)
    titulo = site.findAll('title')
    data = site.findAll('pubdate')
    link = site.findAll('description')

    noticias = []
    noticia = 0 
    
    while True:
        try:
            d = str(data[noticia])
        except:
            break
        
        d = d.split(',')
        d = d[1].split('GMT')
        d = d[0].strip()
        d = d.split(' ')

        for m in meses_a:
            if m == d[1]:
                n = meses_a.index(m)        
                break
            
        data_final = d[0],meses[n],d[2]
        data_final = str(data_final).replace('(','').replace(')','').replace("'",'').replace(' ','').replace(',','').strip()
        
        
        noticia += 1

        t = str(titulo[noticia])
        t = t.split('>')
        t = t[1].split('<')
        t = t[0].split('-')

        titulo_final = t[0].strip().strip()
    
        veiculo_final = t[1].strip().strip()
        

        l = str(link[noticia])
        l = l.split('href=')
        l = l[1].split('"')

        link_final = l[1].strip().strip()
        
        final = {'titulo' : titulo_final, 'veiculo' : veiculo_final, 'data' : data_final, 'link' : link_final}
        noticias.append(final)
        
    return noticias
    

            
