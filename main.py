import requests
from lxml import html
import re
import xlwt

wb = xlwt.Workbook()
ws = wb.add_sheet("test_sheet", cell_overwrite_ok=True)
wb1 = xlwt.Workbook()
ws1 = wb1.add_sheet("test_sheet", cell_overwrite_ok=True)
wb2 = xlwt.Workbook()
ws2 = wb2.add_sheet("test_sheet", cell_overwrite_ok=True)
xl_index = 0
xl_index1 = 0
xl_index2 = 0
url = "https://www.memoiredeshommes.sga.defense.gouv.fr/en/arkotheque/client/mdh/militaires_decedes_seconde_guerre_mondiale/index.php"

data = {
    'action': '1',
    'todo': 'rechercher',
    'le_id': '',
    'multisite': '',
    'r_c_nom': '',
    'r_c_nom_like': '',
    'r_c_prenom': '',
    'r_c_prenom_like': '1',
    'r_c_naissance_jour_mois_annee_jj_debut': '',
    'r_c_naissance_jour_mois_annee_mm_debut': '',
    'r_c_naissance_jour_mois_annee_yyyy_debut': '',
    'r_c_naissance_jour_mois_annee_jj_fin': '',
    'r_c_naissance_jour_mois_annee_mm_fin': '',
    'r_c_naissance_jour_mois_annee_yyyy_fin': '',
    'r_c_id_naissance_departement': '',
    'hidden_c_id_naissance_departement': '',
    'r_c_id_naissance_pays': '',
    'hidden_c_id_naissance_pays': '',
    'r_c_id_mention': '',
    'r_c_pseudonyme': '',
    'r_c_pseudonyme_like': '1',
    'r_c_id_grade': '',
    'hidden_c_id_grade': '',
    'r_c_id_unite': '',
    'hidden_c_id_unite': '',
    'r_c_classe': '',
    'r_c_id_recrutement_bureau': '',
    'hidden_c_id_recrutement_bureau': '',
    'r_c_recrutement_matricule': '',
    'r_c_id_naissance_lieu': '',
    'hidden_c_id_naissance_lieu': '',
    'r_c_deces_jour_mois_annee_jj_debut': '',
    'r_c_deces_jour_mois_annee_mm_debut': '',
    'r_c_deces_jour_mois_annee_yyyy_debut': '',
    'r_c_deces_jour_mois_annee_jj_fin': '',
    'r_c_deces_jour_mois_annee_mm_fin': '',
    'r_c_deces_jour_mois_annee_yyyy_fin': '',
    'r_c_id_deces_lieu': '',
    'hidden_c_id_deces_lieu': '',
    'r_c_deces_lieu_complement': '',
    'r_c_deces_lieu_complement_like': '1',
    'r_c_id_deces_departement': '',
    'hidden_c_id_deces_departement': '',
    'r_c_id_deces_pays': '',
    'hidden_c_id_deces_pays': '',
    'r_c_id_deces_cause': '',
    'r_c_id_deces_cause_like': '1',
    'r_c_id_transcription_etablissement_lieu': '',
    'hidden_c_id_transcription_etablissement_lieu': '',
    'r_c_id_transcription_etablissement_departement': '',
    'hidden_c_id_transcription_etablissement_departement': '',
    'r_c_id_transcription_etablissement_pays': '',
    'hidden_c_id_transcription_etablissement_pays': '',
}
inq = 1
month = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
days = [11, 20, 30]
dedl = 1
add = 'http://www.memoiredeshommes.sga.defense.gouv.fr/en/arkotheque/client/mdh/militaires_decedes_seconde_guerre_mondiale/'
s = requests.Session()

user_agent = {'User-agent': 'Mozilla/5.0'}
for y in range(1860, 1944, 1):
    for m in month:
        d = 1
        if(m=='01'):
            days[2] ='31'
        elif (m == '02'):
            days[2] = '29'
        elif (m == '03'):
            days[2] = '31'
        elif (m == '04'):
            days[2] = '30'
        elif (m == '05'):
            days[2] = '31'
        elif (m == '06'):
            days[2] = '30'
        elif (m == '07'):
            days[2] = '31'
        elif (m == '08'):
            days[2] = '31'
        elif (m == '09'):
            days[2] = '30'
        elif (m == '10'):
            days[2] = '31'
        elif (m == '11'):
            days[2] = '30'
        elif (m == '12'):
            days[2] = '31'


        for day in days:
            d1 = d
            d += 10
            data['r_c_naissance_jour_mois_annee_yyyy_debut'] = '{}'.format(y)
            data['r_c_naissance_jour_mois_annee_mm_debut'] = '{}'.format(m)
            data['r_c_naissance_jour_mois_annee_jj_debut'] = '{}'.format(d1)
            search_crt = str(y)+"/"+str(m)+"/"+str(d1)+"-"+str(y)+"/"+str(m)+"/"+str(day)

            data['r_c_naissance_jour_mois_annee_yyyy_fin'] = '{}'.format(y)
            data['r_c_naissance_jour_mois_annee_mm_fin'] = '{}'.format(m)
            data['r_c_naissance_jour_mois_annee_jj_fin'] = '{}'.format(day)
            d += 1
            p = s.post(url, data=data, headers=user_agent)
            body = html.fromstring(p.text)
            all = 'http://www.memoiredeshommes.sga.defense.gouv.fr/en/arkotheque/client/mdh/militaires_decedes_seconde_guerre_mondiale/resus_rech.php?&aff_tous=1'
            r = s.get(all, headers=user_agent)
            body = html.fromstring(r.text)
            total_res = body.xpath('//*[@id="abecedaire"]//text()')

            total_res = str(total_res).split('on')
            try:
                total_res = total_res[1].split(':')[0]
                print(total_res)
            except IndexError:
                print('vodka error', r.url)
                continue

            for link in body.xpath("//span[contains(@class, 'fiche_detail')]//a//@href"):
                http = add + link

                r = s.get(http, headers=user_agent)
                body = html.fromstring(r.text)
                search_crt = search_crt

                name = body.xpath('/html/body/div/div[3]/div/div[2]/div/form/h1//text()')
                name = str(name).split(' ')
                nom = ''
                prenom = ''
                for n in name:
                    if (str(n).isupper() == True):
                        nom += str(n)
                    else:
                        prenom += str(n)
                ref = str(link).split('ref=')[1]
                ref = ref.replace('&debut=0', '')
                data_naisens = body.xpath('/html/body/div/div[3]/div/div[2]/div/form/h4//text()')
                data_naisens = str(data_naisens).replace("['Né(e) le/en ", '')
                data_naisens = data_naisens.split(' ')[0]
                depa_naisens = body.xpath('/html/body/div/div[3]/div/div[2]/div/form/h4//text()')
                depa_naisens = str(depa_naisens).replace("['Né(e) le/en ", '').replace(data_naisens, '')
                data_deces = body.xpath('/html/body/div/div[3]/div/div[2]/div/form/h3//text()')
                lieu = body.xpath('/html/body/div/div[3]/div/div[2]/div/form/h3//text()')
                if (len(data_deces) != 0):
                    m = re.search("\d", str(data_deces))
                    delete = str(data_deces)[0:m.start()]
                    data_deces = str(data_deces).replace(delete, '')
                    space = (data_deces.find(' '))
                    lieu = data_deces[space:]
                    data_deces = data_deces[0:space]
                statut = ''
                unit = ''
                mention = ''
                cause_deses = ''
                sources = ''
                doc_ref = ''
                for item in body.xpath("//div[contains(@class, 'champ_formulaire')]"):
                    val = (item.xpath('label//text()'))
                    val = str(val).replace("']", "").replace("['", '')

                    if (val == "Status"):
                        statut = item.xpath('span//text()')
                    if (val == "Unit"):
                        unit = item.xpath('span//text()')
                    if (val == "Reference"):
                        mention = item.xpath('span//text()')
                    if (val == "Cause of death"):
                        cause_deses = item.xpath('span//text()')
                    if (val == "Sources"):
                        sources = item.xpath('span//text()')
                    if (val == "Document reference"):
                        doc_ref = item.xpath('span//text()')
                ws.write(xl_index, 0, inq)
                ws.write(xl_index, 1, str(search_crt))
                ws.write(xl_index, 2, total_res)
                ws.write(xl_index, 3, nom)
                ws.write(xl_index, 4, prenom)
                ws.write(xl_index, 5, ref)
                ws.write(xl_index, 6, http)
                ws.write(xl_index, 7, data_naisens)
                ws.write(xl_index, 8, depa_naisens)
                ws.write(xl_index, 9, data_deces)
                ws.write(xl_index, 10, lieu)
                ws.write(xl_index, 11, statut)
                ws.write(xl_index, 12, unit)
                ws.write(xl_index, 13, mention)
                ws.write(xl_index, 14, cause_deses)
                ws.write(xl_index, 15, sources)
                ws.write(xl_index, 16, doc_ref)
                inq += 1
                xl_index += 1
            wb.save('result.xlsx')
            print('saved!!!')
