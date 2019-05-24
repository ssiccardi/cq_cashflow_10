# -*- encoding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2018
#    Stefano Siccardi creativiquadrati snc
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################

from odoo import api, fields, models, _
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from cStringIO import StringIO
import base64

try:
    import xlwt
except ImportError:
    xlwt = None

## HOWTO fix the Unicode special characters issue
## https://stackoverflow.com/questions/21129020/how-to-fix-unicodedecodeerror-ascii-codec-cant-decode-byte
import sys
reload(sys)
sys.setdefaultencoding('utf8')

class PrevisioneInOut(models.TransientModel):
    _name = "previsione.in.out"    
    _description = "Report incassi e pagamenti"
    
    file = fields.Binary('Download', readonly=True)
    data_inizio = fields.Date('Data Saldo', required=True, readonly=True, default=date.today())
    state = fields.Selection([('choose','choose'),('get','get')], default='choose')
    filename = fields.Char('Nome Download', readonly=True, default='Cashflow.xls')
    saldo_tot = fields.Float('Saldo totale')

    def prepare_report(self):

        if not self.saldo_tot:
            saldo_tot=round(0.0,2)
        else:
            saldo_tot=round(self.saldo_tot,2)   
        
        invoice_pool=self.env['account.invoice']
        sale_order_pool=self.env['sale.order']
        purchase_order_pool=self.env['purchase.order']
        righe=[]
        date_start=datetime.now().replace(day=1,hour=0,minute=0,second=0,microsecond=0)
        array_inizi=[date_start]
        array_fine=[date_start+relativedelta(months=1)+relativedelta(days=-1)]
        for el in range(1,5):
            new_inizio=date_start+relativedelta(months=el)
            new_fine=date_start+relativedelta(months=(el+1))+relativedelta(days=-1)
            array_inizi.append(new_inizio)
            array_fine.append(new_fine)
             
        #intestazione incassi
        array_inizio_str=[]
        array_fine_str=[]
        for el in array_inizi:
            inizio_str=datetime.strftime(el, '%d-%m-%y')
            array_inizio_str.append(inizio_str)
        for el2 in array_fine:
            fine_str=datetime.strftime(el2, '%d-%m-%y')
            array_fine_str.append(fine_str)            
       
        testa_mesi = [array_inizio_str[0][3:5]+'/'+array_inizio_str[0][6:],'',array_inizio_str[1][3:5]+'/'+array_inizio_str[1][6:],'',array_inizio_str[2][3:5]+'/'+array_inizio_str[2][6:],'',array_inizio_str[3][3:5]+'/'+array_inizio_str[3][6:],'',array_inizio_str[4][3:5]+'/'+array_inizio_str[4][6:]]
        tot_tot_impo1=0.0
        tot_tot_impo2=0.0
        tot_tot_impo3=0.0
        tot_tot_impo4=0.0        
        tot_tot_impo5=0.0       
### FATTURE CLIENTI, ANCHE QUELLE IN BOZZA 
        #prima riga: mesi delle fatture
        righe.append(['INCASSI-FATTURE','','']+testa_mesi+['',''])
        #seconda riga: intestazione vera e propria
        righe.append(['Cliente','Data Fattura','Numero Fattura','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo',''])
        
        ##giro sulle fatture clienti per cercare quelle che hanno scadenze nei mesi successivi
                
        invoices=invoice_pool.search([('state','not in',('cancel','paid')),('type','in',('out_invoice','out_refund'))], order='date_invoice asc')
        
        tot_impo1=0.0
        tot_impo2=0.0
        tot_impo3=0.0
        tot_impo4=0.0
        tot_impo5=0.0
        for invoice in invoices:
            #sono liste perchè possono capitare due rate all'interno di uno stesso mese
            scad1=[]
            scad2=[]
            scad3=[]
            scad4=[]
            scad5=[]
            impo1=[]
            impo2=[]
            impo3=[]
            impo4=[]
            impo5=[]
            
            data_fattura=invoice.date_invoice or invoice.data_validazione or ''
            pp = ''
            if invoice.partner_id.name:
                pp = invoice.partner_id.name.encode('UTF-8','ignore')            
            if invoice.state!='open':
                totale_cons=invoice.amount_total
            else:
                totale_cons=invoice.residual
            if not invoice.payment_term_id or not invoice.payment_term_id.line_ids:
                #se la fattura non ha data di scadenza,ne data di validazione , ne data di validazione prevista, la salto
                if (not invoice.date_due and not invoice.data_validazione and not invoice.date_invoice) or not totale_cons:
                    continue
                    
                if totale_cons/invoice.amount_total<0.01 and totale_cons<500.:
                    continue                    
                
                #se c'è la data di scadenza, prendo quella, altrimenti prendo la data della fattura , altrimenti la data di validazione prevista
                if invoice.date_due:
                    new_data_fattura=invoice.date_due
                elif invoice.date_invoice:
                    new_data_fattura=invoice.date_invoice
                else:
                    new_data_fattura=invoice.data_validazione
                           
                scadenza=datetime.strptime(new_data_fattura,'%Y-%m-%d')
                #if scadenza>=array_inizi[0] and scadenza<=array_fine[0]:
                if scadenza<=array_fine[0]: ## hanno richiesto che appaiano tutte le scadenze anche precedenti al mese corrente,
                #   ma nella versione per la rel. 7 lo tolgo per fare le prove
                    if invoice.type=='out_refund':
                        tot_impo1=tot_impo1-totale_cons
                        impo1.append(round(-totale_cons,2))
                    else:
                        tot_impo1=tot_impo1+totale_cons
                        impo1.append(round(totale_cons,2))
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad1.append(scadenza_str[:10])
                if scadenza>=array_inizi[1] and scadenza<=array_fine[1]:
                    if invoice.type=='out_refund':
                        tot_impo2=tot_impo2-totale_cons
                        impo2.append(round(-totale_cons,2))
                    else:
                        tot_impo2=tot_impo2+totale_cons
                        impo2.append(round(totale_cons,2))
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad2.append(scadenza_str[:10])
                if scadenza>=array_inizi[2] and scadenza<=array_fine[2]:
                    if invoice.type=='out_refund':
                        tot_impo3=tot_impo3-totale_cons
                        impo3.append(round(-totale_cons,2))
                    else:
                        tot_impo3=tot_impo3+totale_cons
                        impo3.append(round(totale_cons,2))
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad3.append(scadenza_str[:10])
                if scadenza>=array_inizi[3] and scadenza<=array_fine[3]:
                    if invoice.type=='out_refund':
                        tot_impo4=tot_impo4-totale_cons
                        impo4.append(round(-totale_cons,2))
                    else:
                        tot_impo4=tot_impo4+totale_cons
                        impo4.append(round(totale_cons,2))
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad4.append(scadenza_str[:10])
                if scadenza>=array_inizi[4] and scadenza<=array_fine[4]:
                    if invoice.type=='out_refund':
                        tot_impo5=tot_impo5-totale_cons
                        impo5.append(round(-totale_cons,2))
                    else:
                        impo5.append(round(totale_cons,2))
                        tot_impo5=tot_impo5+totale_cons
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad5.append(scadenza_str[:10])
            else:
                #se la fattura non ha data di validazione, ne data di validazione prevista, la salto
                if (not invoice.data_validazione and not invoice.date_invoice) or not totale_cons:
                    continue
                
                #se c'è la data di validazione, prendo quella, altrimenti prendo la data di validazione prevista
                if invoice.date_invoice:
                    new_data_fattura=invoice.date_invoice
                else:
                    new_data_fattura=invoice.data_validazione
           
                rate=[]
                for line in invoice.payment_term_id.compute(invoice.amount_total, new_data_fattura)[0]:
                    rate.append([line[1],line[0]])
            ###toglie l'ammontare del pagato dalle rate  
                ammontare_pagato=round(invoice.amount_total-totale_cons,2)
                ammontare_rate=0.
                tolto_pagato=False
                for rata in rate:
                    ammontare_rate+=round(rata[0],2)
                    
                    if ammontare_pagato>ammontare_rate:
                        continue
                        
                    if ammontare_pagato and not tolto_pagato:
                        tolto_pagato=True
                        rata[0]=ammontare_rate-ammontare_pagato
                    
                    scadenza=datetime.strptime(rata[1],'%Y-%m-%d')
                    rata=rata[0]
                    
                    if rata/invoice.amount_total<0.01 and rata<500.:
                        continue                    
                        
                    if scadenza<=array_fine[0]: ## hanno richiesto che appaiano tutte le scadenze anche precedenti al mese corrente
                        if invoice.type=='out_refund':
                            tot_impo1=tot_impo1-rata
                            impo1.append(round(-rata,2))
                        else:
                            tot_impo1=tot_impo1+rata
                            impo1.append(round(rata,2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad1.append(scadenza_str[:10])
                    if scadenza>=array_inizi[1] and scadenza<=array_fine[1]:
                        if invoice.type=='out_refund':
                            tot_impo2=tot_impo2-rata
                            impo2.append(round(-rata,2))
                        else:
                            tot_impo2=tot_impo2+rata
                            impo2.append(round(rata,2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad2.append(scadenza_str[:10])
                    if scadenza>=array_inizi[2] and scadenza<=array_fine[2]:
                        if invoice.type=='out_refund':
                            tot_impo3=tot_impo3-rata
                            impo3.append(round(-rata,2))
                        else:
                            tot_impo3=tot_impo3+rata
                            impo3.append(round(rata,2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad3.append(scadenza_str[:10])
                    if scadenza>=array_inizi[3] and scadenza<=array_fine[3]:
                        if invoice.type=='out_refund':
                            tot_impo4=tot_impo4-rata
                            impo4.append(round(-rata,2))
                        else:
                            tot_impo4=tot_impo4+rata
                            impo4.append(round(rata,2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad4.append(scadenza_str[:10])
                    if scadenza>=array_inizi[4] and scadenza<=array_fine[4]:
                        if invoice.type=='out_refund':
                            tot_impo5=tot_impo5-rata
                            impo5.append(round(-rata,2))
                        else:
                            tot_impo5=tot_impo5+rata
                            impo5.append(round(rata,2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad5.append(scadenza_str[:10])
            if impo1 or impo2 or impo3 or impo4 or impo5:
                max_len=max(len(impo1),len(impo2),len(impo3),len(impo4),len(impo5))
                for i in range(max_len):
                    if len(impo1)<max_len:
                        impo1.append('')
                    if len(impo2)<max_len:
                        impo2.append('')
                    if len(impo3)<max_len:
                        impo3.append('')
                    if len(impo4)<max_len:
                        impo4.append('')
                    if len(impo5)<max_len:
                        impo5.append('')
                    if len(scad1)<max_len:
                        scad1.append('')
                    if len(scad2)<max_len:
                        scad2.append('')
                    if len(scad3)<max_len:
                        scad3.append('')
                    if len(scad4)<max_len:
                        scad4.append('')
                    if len(scad5)<max_len:
                        scad5.append('')
                righe.append([pp,data_fattura,invoice.number or 'Bozza',scad1[0],impo1[0],scad2[0],impo2[0],scad3[0],impo3[0],scad4[0],impo4[0],scad5[0],impo5[0],''])
                for i in range(1,max_len):
                    righe.append(['','','',scad1[i],impo1[i],scad2[i],impo2[i],scad3[i],impo3[i],scad4[i],impo4[i],scad5[i],impo5[i],''])
        
        righe.append(['','','','','','','','','','','','','',''])
        righe.append(['TOTALE','','','',round(tot_impo1,2),'',round(tot_impo2,2),'',round(tot_impo3,2),'',round(tot_impo4,2),'',round(tot_impo5,2),''])
        line_total_fatt_v=['INCASSI FATTURE','',round(tot_impo1,2),round(tot_impo2,2),round(tot_impo3,2),round(tot_impo4,2),round(tot_impo5,2),'','','','','','','']          
        righe.append(['','','','','','','','','','','','','',''])
        righe.append(['','','','','','','','','','','','','',''])
        tot_tot_impo1=tot_tot_impo1+round(tot_impo1,2)
        tot_tot_impo2=tot_tot_impo2+round(tot_impo2,2)
        tot_tot_impo3=tot_tot_impo3+round(tot_impo3,2)
        tot_tot_impo4=tot_tot_impo4+round(tot_impo4,2)
        tot_tot_impo5=tot_tot_impo5+round(tot_impo5,2)        
                        
        
### ORDINI CLIENTI        
        #prima riga: mesi degli ordini
        righe.append(['INCASSI-ORDINI','','','']+testa_mesi+[''])
        #seconda riga: intestazione vera e propria
        righe.append(['Cliente','Data Ordine','Numero Ordine','Data Impegno','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo'])        
        
        # ordini di vendita
        sale_orders=sale_order_pool.search([('state','not in',('draft','cancel','done','sent'))],order='date_order asc')        
        tot_impo1=0.0
        tot_impo2=0.0
        tot_impo3=0.0
        tot_impo4=0.0
        tot_impo5=0.0

        for sale_order in sale_orders:
            sn=''
            pp = ''
            if sale_order.partner_id.name:
                pp = sale_order.partner_id.name.encode('UTF-8','ignore')            
            #se non c'è la tabella divisione fatturazione considera l'ordine con una rata intera con data dell'ordine
            div_fatts=self.env['divisione.fatturazione.sale'].search([('order_id','=',sale_order.id)],order='data_prevista asc, importo asc')
            
            if (not sale_order.date_order and not div_fatts) or not sale_order.amount_total:
                continue
                
            if not div_fatts:
                unique=[sale_order.date_order[:10]]
                importo_impegno_array=[sale_order.amount_total]
            else:
                unique=[]    
                impegno_array=[]
                for div_fatt in div_fatts:
                    impegno_array.append(div_fatt.data_prevista)
                [unique.append(item) for item in impegno_array if item not in unique] 
               
                importo_impegno_array=[]
                rate=[]
                for data in unique:             
                    importo_impegno_data=0.0
                    for line in div_fatts:
                        if line.data_prevista and line.data_prevista==data:
                            importo_impegno_data=importo_impegno_data+line.importo
                    importo_impegno_array.append(importo_impegno_data)
                    rate.append([importo_impegno_data,data])                         
            
            #cerco fatture (parziali, bozze) collegate
            fcoll=[]
            for inv in sale_order.invoice_ids:
                if inv.state!='cancel' and inv.amount_total:
                    if not inv.payment_term_id or not inv.payment_term_id.line_ids:
                        if not inv.date_due and not inv.data_validazione and not inv.date_invoice:
                            continue
                    else:
                        if not inv.data_validazione and not inv.date_invoice:
                            continue                    
                    fcoll.append(inv)                   
            
            #se ci sono fatture parziali collegate, vado a considerare l'importo fatturato e lo scalo dalle rate iniziali della tabella "Divisione fatturazione".
            if fcoll:
                unique=[]
                importo_impegno_array=[]
                importo_fatturato=0.                           
                for fco in fcoll:
                    importo_fatturato+=fco.amount_total
                if not div_fatts:
                    if importo_fatturato<sale_order.amount_total:
                        unique=[sale_order.date_order[:10]]
                        importo_impegno_array=[sale_order.amount_total-importo_fatturato]
                    else:
                        unique=[]
                else:                                         
                    ammontare_rate=0.
                    tolto_fatturato=False
                    aggiungere_succ_rata=False
                    importo_da_aggiungere_succ_rata=0.
                    i=0
                    for rata in rate:
                        i+=1
                        ammontare_rate+=rata[0]
                        if importo_fatturato>ammontare_rate:
                            continue
                        if importo_fatturato and not tolto_fatturato:
                            tolto_fatturato=True
                            rata[0]=round(ammontare_rate-importo_fatturato,2)
                            if rata[0]/sale_order.amount_total<=0.05 and i<len(rate): #se l'avanzo della rata è minore del 5% del totale lo aggiugo alla rata successiva,se c'è
                                aggiungere_succ_rata=True
                                importo_da_aggiungere_succ_rata=rata[0]
                                continue
                        if aggiungere_succ_rata and importo_da_aggiungere_succ_rata:
                            rata[0]+=importo_da_aggiungere_succ_rata
                            aggiungere_succ_rata=False
                            importo_da_aggiungere_succ_rata=0.
                        unique.append(rata[1])
                        importo_impegno_array.append(rata[0])
            
            for el in range(0,len(unique)):         
                scad1=[]
                scad2=[]
                scad3=[]
                scad4=[]
                scad5=[]
                impo1=[]
                impo2=[]
                impo3=[]
                impo4=[]
                impo5=[]
                
                if importo_impegno_array[el]/sale_order.amount_total<0.01 and importo_impegno_array[el]<500.:
                    continue
                
                if not sale_order.payment_term_id or not sale_order.payment_term_id.line_ids:
                    scadenza=datetime.strptime(unique[el],'%Y-%m-%d')
                    #if scadenza>=array_inizi[0] and scadenza<=array_fine[0]:
                    if scadenza<=array_fine[0]: ## hanno richiesto che appaiano tutte le scadenze anche precedenti al mese corrente
                        tot_impo1=tot_impo1+importo_impegno_array[el]
                        impo1.append(round(importo_impegno_array[el],2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad1.append(scadenza_str[:10])
                    if scadenza>=array_inizi[1] and scadenza<=array_fine[1]:
                        tot_impo2=tot_impo2+importo_impegno_array[el]
                        impo2.append(round(importo_impegno_array[el],2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad2.append(scadenza_str[:10])
                    if scadenza>=array_inizi[2] and scadenza<=array_fine[2]:
                        tot_impo3=tot_impo3+importo_impegno_array[el]
                        impo3.append(round(importo_impegno_array[el],2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad3.append(scadenza_str[:10])
                    if scadenza>=array_inizi[3] and scadenza<=array_fine[3]:
                        tot_impo4=tot_impo4+importo_impegno_array[el]
                        impo4.append(round(importo_impegno_array[el],2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad4.append(scadenza_str[:10])
                    if scadenza>=array_inizi[4] and scadenza<=array_fine[4]:
                        tot_impo5=tot_impo5+importo_impegno_array[el]
                        impo5.append(round(importo_impegno_array[el],2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad5.append(scadenza_str[:10])
                else:      
                    for line in sale_order.payment_term_id.compute(importo_impegno_array[el], unique[el])[0]:
                        scadenza = scadenza=datetime.strptime(line[0],'%Y-%m-%d')
                        rata = line[1]
                        if rata/sale_order.amount_total<0.01 and rata<500.:
                            continue                            
                
                        #if scadenza>=array_inizi[0] and scadenza<=array_fine[0]:
                        if scadenza<=array_fine[0]: ## hanno richiesto che appaiano tutte le scadenze anche precedenti al mese corrente
                            tot_impo1=tot_impo1+rata
                            impo1.append(round(rata,2))
                            scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                            scad1.append(scadenza_str[:10])
                        if scadenza>=array_inizi[1] and scadenza<=array_fine[1]:
                            tot_impo2=tot_impo2+rata
                            impo2.append(round(rata,2))
                            scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                            scad2.append(scadenza_str[:10])
                        if scadenza>=array_inizi[2] and scadenza<=array_fine[2]:
                            tot_impo3=tot_impo3+rata
                            impo3.append(round(rata,2))
                            scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                            scad3.append(scadenza_str[:10])
                        if scadenza>=array_inizi[3] and scadenza<=array_fine[3]:
                            tot_impo4=tot_impo4+rata
                            impo4.append(round(rata,2))
                            scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                            scad4.append(scadenza_str[:10])
                        if scadenza>=array_inizi[4] and scadenza<=array_fine[4]:
                            tot_impo5=tot_impo5+rata
                            impo5.append(round(rata,2))
                            scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                            scad5.append(scadenza_str[:10])
            
                if impo1 or impo2 or impo3 or impo4 or impo5:
                    max_len=max(len(impo1),len(impo2),len(impo3),len(impo4),len(impo5))
                    for i in range(max_len):
                        if len(impo1)<max_len:
                            impo1.append('')
                        if len(impo2)<max_len:
                            impo2.append('')
                        if len(impo3)<max_len:
                            impo3.append('')
                        if len(impo4)<max_len:
                            impo4.append('')
                        if len(impo5)<max_len:
                            impo5.append('')
                        if len(scad1)<max_len:
                            scad1.append('')
                        if len(scad2)<max_len:
                            scad2.append('')
                        if len(scad3)<max_len:
                            scad3.append('')
                        if len(scad4)<max_len:
                            scad4.append('')
                        if len(scad5)<max_len:
                            scad5.append('')                
                    if sn:
                        righe.append(['','','',unique[el],scad1[0],impo1[0],scad2[0],impo2[0],scad3[0],impo3[0],scad4[0],impo4[0],scad5[0],impo5[0]])
                    else:
                        sn=sale_order.name
                        righe.append([pp,sale_order.date_order[:10],sn,unique[el],scad1[0],impo1[0],scad2[0],impo2[0],scad3[0],impo3[0],scad4[0],impo4[0],scad5[0],impo5[0]])
                    for i in range(1,max_len):
                        righe.append(['','','','',scad1[i],impo1[i],scad2[i],impo2[i],scad3[i],impo3[i],scad4[i],impo4[i],scad5[i],impo5[i]])                        

        righe.append(['','','','','','','','','','','','','',''])
        righe.append(['TOTALE','','','','',round(tot_impo1,2),'',round(tot_impo2,2),'',round(tot_impo3,2),'',round(tot_impo4,2),'',round(tot_impo5,2)])
        line_total_ord_v=['INCASSI ORDINI','',round(tot_impo1,2),round(tot_impo2,2),round(tot_impo3,2),round(tot_impo4,2),round(tot_impo5,2),'','','','','','','']             
        righe.append(['','','','','','','','','','','','','',''])
        righe.append(['','','','','','','','','','','','','',''])

        tot_tot_impo1=tot_tot_impo1+round(tot_impo1,2)
        tot_tot_impo2=tot_tot_impo2+round(tot_impo2,2)
        tot_tot_impo3=tot_tot_impo3+round(tot_impo3,2)
        tot_tot_impo4=tot_tot_impo4+round(tot_impo4,2)
        tot_tot_impo5=tot_tot_impo5+round(tot_impo5,2)


### FATTURE FORNITORI
        #prima riga: mesi delle fatture
        righe.append(['PAGAMENTI-FATTURE','','']+testa_mesi+['',''])
        #seconda riga: intestazione vera e propria
        righe.append(['Fornitori','Data Fattura','Numero Fattura','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo',''])
        
        ##giro sulle fatture fornitori per cercare quelle che hanno scadenze nei mesi successivi                
        invoices=invoice_pool.search([('state','not in',('cancel','paid')),('type','in',('in_invoice','in_refund'))], order='date_invoice asc')
        
        tot_impo1=0.0
        tot_impo2=0.0
        tot_impo3=0.0
        tot_impo4=0.0
        tot_impo5=0.0
        for invoice in invoices:
            scad1=[]
            scad2=[]
            scad3=[]
            scad4=[]
            scad5=[]
            impo1=[]
            impo2=[]
            impo3=[]
            impo4=[]
            impo5=[]
               
            data_fattura=invoice.date_invoice or invoice.data_validazione or ''
            pp = ''
            if invoice.partner_id.name:
                pp = invoice.partner_id.name.encode('UTF-8','ignore')            
            if invoice.state!='open':
                totale_cons=invoice.amount_total
            else:
                totale_cons=invoice.residual            
            if not invoice.payment_term_id or not invoice.payment_term_id.line_ids:
                
                #se la fattura non ha data di scadenza e nemmeno data di validazione, la salto
                if (not invoice.date_invoice and not invoice.data_validazione and not invoice.date_due) or not totale_cons:
                    continue
                    
                if totale_cons/invoice.amount_total<0.01 and totale_cons<500.:
                    continue

                #se c'è la data di scadenza, prendo quella, altrimenti prendo la data della fattura , altrimenti la data di validazione prevista
                if invoice.date_due:
                    new_data_fattura=invoice.date_due
                elif invoice.date_invoice:
                    new_data_fattura=invoice.date_invoice
                else:
                    new_data_fattura=invoice.data_validazione
                            
                scadenza=datetime.strptime(new_data_fattura,'%Y-%m-%d')
                #if scadenza>=array_inizi[0] and scadenza<=array_fine[0]:
                if scadenza<=array_fine[0]: ## hanno richiesto che appaiano tutte le scadenze anche precedenti al mese corrente
                    if invoice.type=='in_refund':
                        impo1.append(round(-totale_cons,2))
                        tot_impo1=tot_impo1-totale_cons
                    else:
                        impo1.append(round(totale_cons,2))
                        tot_impo1=tot_impo1+totale_cons
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad1.append(scadenza_str[:10])
                if scadenza>=array_inizi[1] and scadenza<=array_fine[1]:
                    if invoice.type=='in_refund':
                        impo2.append(round(-totale_cons,2))
                        tot_impo2=tot_impo2-totale_cons
                    else:
                        impo2.append(round(totale_cons,2))
                        tot_impo2=tot_impo2+totale_cons
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad2.append(scadenza_str[:10])
                if scadenza>=array_inizi[2] and scadenza<=array_fine[2]:
                    if invoice.type=='in_refund':
                        impo3.append(round(-totale_cons,2))
                        tot_impo3=tot_impo3-totale_cons
                    else:
                        impo3.append(round(totale_cons,2))
                        tot_impo3=tot_impo3+totale_cons
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad3.append(scadenza_str[:10])
                if scadenza>=array_inizi[3] and scadenza<=array_fine[3]:
                    if invoice.type=='in_refund':
                        impo4.append(round(-totale_cons,2))
                        tot_impo4=tot_impo4-totale_cons
                    else:
                        impo4.append(round(totale_cons,2))
                        tot_impo4=tot_impo4+totale_cons
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad4.append(scadenza_str[:10])
                if scadenza>=array_inizi[4] and scadenza<=array_fine[4]:
                    if invoice.type=='in_refund':
                        impo5.append(round(-totale_cons,2))
                        tot_impo5=tot_impo5-totale_cons
                    else:
                        impo5.append(round(totale_cons,2))
                        tot_impo5=tot_impo5+totale_cons
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad5.append(scadenza_str[:10])
            else:
                #se la fattura non ha data di validazione, ne data di validazione prevista, la salto
                if (not invoice.data_validazione and not invoice.date_invoice) or not totale_cons:
                    continue
                
                #se c'è la data di validazione, prendo quella, altrimenti prendo la data di validazione prevista
                if invoice.date_invoice:
                    new_data_fattura=invoice.date_invoice
                else:
                    new_data_fattura=invoice.data_validazione
           
                rate=[]
                for line in invoice.payment_term_id.compute(invoice.amount_total, new_data_fattura)[0]:
                    rate.append([line[1],line[0]])
                ###toglie l'ammontare del pagato dalle rate
                ammontare_pagato=round(invoice.amount_total-totale_cons,2)
                ammontare_rate=0.
                tolto_pagato=False
                for rata in rate:
                    ammontare_rate+=round(rata[0],2)
                    
                    if ammontare_pagato>ammontare_rate:
                        continue
                        
                    if ammontare_pagato and not tolto_pagato:
                        tolto_pagato=True
                        rata[0]=ammontare_rate-ammontare_pagato
                    
                    scadenza = datetime.strptime(rata[1],'%Y-%m-%d')
                    rata=rata[0]

                    if rata/invoice.amount_total<0.01 and rata<500.:
                        continue                    
                    
                    #if scadenza>=array_inizi[0] and scadenza<=array_fine[0]:
                    if scadenza<=array_fine[0]: ## hanno richiesto che appaiano tutte le scadenze anche precedenti al mese corrente
                        if invoice.type=='in_refund':
                            tot_impo1=tot_impo1-rata
                            impo1.append(round(-rata,2))
                        else:
                            tot_impo1=tot_impo1+rata
                            impo1.append(round(rata,2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad1.append(scadenza_str[:10])
                    if scadenza>=array_inizi[1] and scadenza<=array_fine[1]:
                        if invoice.type=='in_refund':
                            tot_impo2=tot_impo2-rata
                            impo2.append(round(-rata,2))
                        else:
                            tot_impo2=tot_impo2+rata
                            impo2.append(round(rata,2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad2.append(scadenza_str[:10])
                    if scadenza>=array_inizi[2] and scadenza<=array_fine[2]:
                        if invoice.type=='in_refund':
                            tot_impo3=tot_impo3-rata
                            impo3.append(round(-rata,2))
                        else:
                            tot_impo3=tot_impo3+rata
                            impo3.append(round(rata,2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad3.append(scadenza_str[:10])
                    if scadenza>=array_inizi[3] and scadenza<=array_fine[3]:
                        if invoice.type=='in_refund':
                            tot_impo4=tot_impo4-rata
                            impo4.append(round(-rata,2))
                        else:
                            tot_impo4=tot_impo4+rata
                            impo4.append(round(rata,2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad4.append(scadenza_str[:10])
                    if scadenza>=array_inizi[4] and scadenza<=array_fine[4]:
                        if invoice.type=='in_refund':
                            tot_impo5=tot_impo5-rata
                            impo5.append(round(-rata,2))
                        else:
                            tot_impo5=tot_impo5+rata
                            impo5.append(round(rata,2))
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad5.append(scadenza_str[:10])
            
            if impo1 or impo2 or impo3 or impo4 or impo5:
                max_len=max(len(impo1),len(impo2),len(impo3),len(impo4),len(impo5))
                for i in range(max_len):
                    if len(impo1)<max_len:
                        impo1.append('')
                    if len(impo2)<max_len:
                        impo2.append('')
                    if len(impo3)<max_len:
                        impo3.append('')
                    if len(impo4)<max_len:
                        impo4.append('')
                    if len(impo5)<max_len:
                        impo5.append('')
                    if len(scad1)<max_len:
                        scad1.append('')
                    if len(scad2)<max_len:
                        scad2.append('')
                    if len(scad3)<max_len:
                        scad3.append('')
                    if len(scad4)<max_len:
                        scad4.append('')
                    if len(scad5)<max_len:
                        scad5.append('')
                righe.append([pp,data_fattura,invoice.number or 'Bozza',scad1[0],impo1[0],scad2[0],impo2[0],scad3[0],impo3[0],scad4[0],impo4[0],scad5[0],impo5[0],''])
                for i in range(1,max_len):
                    righe.append(['','','',scad1[i],impo1[i],scad2[i],impo2[i],scad3[i],impo3[i],scad4[i],impo4[i],scad5[i],impo5[i],''])

        righe.append(['','','','','','','','','','','','','',''])
        righe.append(['TOTALE','','','',round(tot_impo1,2),'',round(tot_impo2,2),'',round(tot_impo3,2),'',round(tot_impo4,2),'',round(tot_impo5,2),''])
        line_total_fatt_a=['PAGAMENTI FATTURE','',-1*round(tot_impo1,2),-1*round(tot_impo2,2),-1*round(tot_impo3,2),-1*round(tot_impo4,2),-1*round(tot_impo5,2),'','','','','','','']        
        righe.append(['','','','','','','','','','','','','',''])
        righe.append(['','','','','','','','','','','','','',''])

        tot_tot_impo1=tot_tot_impo1-round(tot_impo1,2)
        tot_tot_impo2=tot_tot_impo2-round(tot_impo2,2)
        tot_tot_impo3=tot_tot_impo3-round(tot_impo3,2)
        tot_tot_impo4=tot_tot_impo4-round(tot_impo4,2)
        tot_tot_impo5=tot_tot_impo5-round(tot_impo5,2)


### ORDINI FORNITORI        
        #prima riga: mesi degli ordini
        righe.append(['PAGAMENTI-ORDINI','','','']+testa_mesi+[''])
        #seconda riga: intestazione vera e propria
        righe.append(['Fornitori','Data Ordine','Numero Ordine','Data preventivata','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo'])        
        
        # ordini di acquisto
        purchase_orders=purchase_order_pool.search([('state','not in',('draft','cancel','done','sent'))],order='date_order asc')
        tot_impo1=0.0
        tot_impo2=0.0
        tot_impo3=0.0
        tot_impo4=0.0
        tot_impo5=0.0

        for purchase_order in purchase_orders:
            pn=''
            pp = ''
            if purchase_order.partner_id.name:
                pp = purchase_order.partner_id.name.encode('UTF-8','ignore')              
            #se non c'è la tabella divisione fatturazione considera l'ordine con una rata intera con data dell'ordine
            div_fatts=self.env['divisione.fatturazione.purchase'].search([('order_id','=',purchase_order.id)],order='data_prevista asc, importo asc')
            
            if (not purchase_order.date_order and not div_fatts) or not purchase_order.amount_total:
                continue
                
            if not div_fatts:
                unique=[purchase_order.date_order[:10]]    
                importo_impegno_array=[purchase_order.amount_total]
            else:
                unique=[]    
                impegno_array=[]
                for div_fatt in div_fatts:
                    impegno_array.append(div_fatt.data_prevista)
                [unique.append(item) for item in impegno_array if item not in unique] 
               
                importo_impegno_array=[]
                rate=[]
                for data in unique:             
                    importo_impegno_data=0.0
                    for line in div_fatts:
                        if line.data_prevista and line.data_prevista==data:
                            importo_impegno_data=importo_impegno_data+line.importo
                    importo_impegno_array.append(importo_impegno_data)
                    rate.append([importo_impegno_data,data])
            
            fcoll = []
            for inv in purchase_order.invoice_ids:
                if inv.state!='cancel' and inv.amount_total:
                    if not inv.payment_term_id or not inv.payment_term_id.line_ids:
                        if not inv.date_due and not inv.data_validazione and not inv.date_invoice:
                            continue
                    else:
                        if not inv.data_validazione and not inv.date_invoice:
                            continue                    
                    fcoll.append(inv) 
            
            #se ci sono fatture parziali collegate, vado a considerare l'importo fatturato e lo scalo dalle rate iniziali della tabella "Divisione fatturazione".
            if fcoll:
                unique=[]
                importo_impegno_array=[]
                importo_fatturato=0.                           
                for fco in fcoll:
                    importo_fatturato+=fco.amount_total
                if not div_fatts:
                    unique=[purchase_order.date_order[:10]]    
                    importo_impegno_array=[purchase_order.amount_total-importo_fatturato]
                else:                                             
                    ammontare_rate=0.
                    tolto_fatturato=False
                    aggiungere_succ_rata=False
                    importo_da_aggiungere_succ_rata=0.
                    i=0
                    for rata in rate:
                        i+=1
                        ammontare_rate+=rata[0]
                        if importo_fatturato>ammontare_rate:
                            continue
                        if importo_fatturato and not tolto_fatturato:
                            tolto_fatturato=True
                            rata[0]=round(ammontare_rate-importo_fatturato,2)
                            if rata[0]/purchase_order.amount_total<=0.05 and i<len(rate): #se l'avanzo della rata è minore del 5% del totale lo aggiugo alla rata successiva,se c'è
                                aggiungere_succ_rata=True
                                importo_da_aggiungere_succ_rata=rata[0]
                                continue
                        if aggiungere_succ_rata and importo_da_aggiungere_succ_rata:
                            rata[0]+=importo_da_aggiungere_succ_rata
                            aggiungere_succ_rata=False
                            importo_da_aggiungere_succ_rata=0.
                        unique.append(rata[1])
                        importo_impegno_array.append(rata[0])            

            for el in range(0,len(unique)):         
                scad1=[]
                scad2=[]
                scad3=[]
                scad4=[]
                scad5=[]
                impo1=[]
                impo2=[]
                impo3=[]
                impo4=[]
                impo5=[]      
                
                if importo_impegno_array[el]/purchase_order.amount_total<0.01 and importo_impegno_array[el]<500.:
                    continue                

                if not purchase_order.payment_term_id or not purchase_order.payment_term_id.line_ids:
                    scadenza=datetime.strptime(unique[el],'%Y-%m-%d')
                    #if scadenza>=array_inizi[0] and scadenza<=array_fine[0]:
                    if scadenza<=array_fine[0]: ## hanno richiesto che appaiano tutte le scadenze anche precedenti al mese corrente
                        tot_impo1=tot_impo1+importo_impegno_array[el]
                        impo1.append(round(importo_impegno_array[el],2))                    
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad1.append(scadenza_str[:10])
                    if scadenza>=array_inizi[1] and scadenza<=array_fine[1]:
                        tot_impo2=tot_impo2+importo_impegno_array[el]
                        impo2.append(round(importo_impegno_array[el],2))  
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad2.append(scadenza_str[:10])
                    if scadenza>=array_inizi[2] and scadenza<=array_fine[2]:
                        tot_impo3=tot_impo3+importo_impegno_array[el]
                        impo3.append(round(importo_impegno_array[el],2))  
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad3.append(scadenza_str[:10])
                    if scadenza>=array_inizi[3] and scadenza<=array_fine[3]:
                        tot_impo4=tot_impo4+importo_impegno_array[el]
                        impo4.append(round(importo_impegno_array[el],2))  
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad4.append(scadenza_str[:10])
                    if scadenza>=array_inizi[4] and scadenza<=array_fine[4]:
                        tot_impo5=tot_impo5+importo_impegno_array[el]
                        impo5.append(round(importo_impegno_array[el],2))  
                        scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                        scad5.append(scadenza_str[:10])
                else:
                    for line in purchase_order.payment_term_id.compute(importo_impegno_array[el], unique[el])[0]:
                        scadenza = datetime.strptime(line[0],'%Y-%m-%d')
                        rata = line[1]
                            
                        if rata/purchase_order.amount_total<0.01 and rata<500.:
                            continue                            
                
                        #if scadenza>=array_inizi[0] and scadenza<=array_fine[0]:
                        if scadenza<=array_fine[0]: ## hanno richiesto che appaiano tutte le scadenze anche precedenti al mese corrente
                            tot_impo1=tot_impo1+rata
                            impo1.append(round(rata,2))
                            scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                            scad1.append(scadenza_str[:10])
                        if scadenza>=array_inizi[1] and scadenza<=array_fine[1]:
                            tot_impo2=tot_impo2+rata
                            impo2.append(round(rata,2))
                            scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                            scad2.append(scadenza_str[:10])
                        if scadenza>=array_inizi[2] and scadenza<=array_fine[2]:
                            tot_impo3=tot_impo3+rata
                            impo3.append(round(rata,2))
                            scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                            scad3.append(scadenza_str[:10])
                        if scadenza>=array_inizi[3] and scadenza<=array_fine[3]:
                            tot_impo4=tot_impo4+rata
                            impo4.append(round(rata,2))
                            scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                            scad4.append(scadenza_str[:10])
                        if scadenza>=array_inizi[4] and scadenza<=array_fine[4]:
                            tot_impo5=tot_impo5+rata
                            impo5.append(round(rata,2))
                            scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                            scad5.append(scadenza_str[:10])
        
                if impo1 or impo2 or impo3 or impo4 or impo5:
                    max_len=max(len(impo1),len(impo2),len(impo3),len(impo4),len(impo5))
                    for i in range(max_len):
                        if len(impo1)<max_len:
                            impo1.append('')
                        if len(impo2)<max_len:
                            impo2.append('')
                        if len(impo3)<max_len:
                            impo3.append('')
                        if len(impo4)<max_len:
                            impo4.append('')
                        if len(impo5)<max_len:
                            impo5.append('')
                        if len(scad1)<max_len:
                            scad1.append('')
                        if len(scad2)<max_len:
                            scad2.append('')
                        if len(scad3)<max_len:
                            scad3.append('')
                        if len(scad4)<max_len:
                            scad4.append('')
                        if len(scad5)<max_len:
                            scad5.append('')                
                    if pn:
                        righe.append(['','','',unique[el],scad1[0],impo1[0],scad2[0],impo2[0],scad3[0],impo3[0],scad4[0],impo4[0],scad5[0],impo5[0]])
                    else:
                        pn=purchase_order.name
                        righe.append([pp,purchase_order.date_order[:10],pn,unique[el],scad1[0],impo1[0],scad2[0],impo2[0],scad3[0],impo3[0],scad4[0],impo4[0],scad5[0],impo5[0]])
                    for i in range(1,max_len):
                        righe.append(['','','','',scad1[i],impo1[i],scad2[i],impo2[i],scad3[i],impo3[i],scad4[i],impo4[i],scad5[i],impo5[i]])       

        righe.append(['','','','','','','','','','','','','',''])
        righe.append(['TOTALE','','','','',round(tot_impo1,2),'',round(tot_impo2,2),'',round(tot_impo3,2),'',round(tot_impo4,2),'',round(tot_impo5,2)])
        line_total_ord_a=['PAGAMENTI ORDINI','',-1*round(tot_impo1,2),-1*round(tot_impo2,2),-1*round(tot_impo3,2),-1*round(tot_impo4,2),-1*round(tot_impo5,2),'','','','','','','']
        righe.append(['','','','','','','','','','','','','',''])
        righe.append(['','','','','','','','','','','','','',''])
        tot_tot_impo1=tot_tot_impo1-round(tot_impo1,2)
        tot_tot_impo2=tot_tot_impo2-round(tot_impo2,2)
        tot_tot_impo3=tot_tot_impo3-round(tot_impo3,2)
        tot_tot_impo4=tot_tot_impo4-round(tot_impo4,2)
        tot_tot_impo5=tot_tot_impo5-round(tot_impo5,2)


###ALTRE SCADENZE
    #prendo tutti i conti nelle varie configurazioni del cashflow
        sql="SELECT account_id,type FROM config_cashflow_base_line ORDER BY conf_id asc"
        self._cr.execute(sql)
        account_ids=self._cr.fetchall()
            
        conti={}
        for account_id in account_ids:
            conti.update({account_id[0]:account_id[1]})

        #prima riga: mesi scadenze
        righe.append(['ALTRE SCADENZE','','','']+testa_mesi+[''])
        #seconda riga: intestazione vera e propria
        righe.append(['Partner','Nome','Registrazione contabile','Data effettiva','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo','Data Scadenza','Saldo']) 
        tot_impo1=0.0
        tot_impo2=0.0
        tot_impo3=0.0
        tot_impo4=0.0
        tot_impo5=0.0        
        if conti:
            move_lines=self.env['account.move.line'].search([('account_id','in',tuple(conti.keys()))],order='create_date asc')
            for move_line in move_lines:
                scad1=''
                scad2=''
                scad3=''
                scad4=''
                scad5=''
                impo1=''
                impo2=''
                impo3=''
                impo4=''
                impo5='' 

                if move_line.amount_residual == 0.0:
                    continue
                else:
                    new_imp=move_line.amount_residual
                    if conti[move_line.account_id.id]=='costo':
                        segno=-1
                    else:
                        segno=1
                
                if move_line.date_maturity:
                    scadenza=datetime.strptime(move_line.date_maturity,'%Y-%m-%d')
                elif move_line.move_id:
                    scadenza=datetime.strptime(move_line.move_id.date,'%Y-%m-%d')
                else:
                    continue
                if scadenza<=array_fine[0]: ## hanno richiesto che appaiano tutte le scadenze anche precedenti al mese corrente
                    impo1=new_imp*segno
                    tot_impo1=tot_impo1+impo1                        
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad1=scadenza_str[:10]
                if scadenza>=array_inizi[1] and scadenza<=array_fine[1]:
                    impo2=new_imp*segno
                    tot_impo2=tot_impo2+impo2
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad2=scadenza_str[:10]                    
                if scadenza>=array_inizi[2] and scadenza<=array_fine[2]:
                    impo3=new_imp*segno
                    tot_impo3=tot_impo3+impo3
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad3=scadenza_str[:10]                
                if scadenza>=array_inizi[3] and scadenza<=array_fine[3]:
                    impo4=new_imp*segno
                    tot_impo4=tot_impo4+impo4
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad4=scadenza_str[:10]
                if scadenza>=array_inizi[4] and scadenza<=array_fine[4]:
                    impo5=new_imp*segno
                    tot_impo5=tot_impo5+impo5
                    scadenza_str=datetime.strftime(scadenza,'%Y-%m-%d %H:%M:%S')
                    scad5=scadenza_str[:10] 


                if impo1<>'' or impo2<>'' or impo3<>'' or impo4<>'' or impo5<>'':
            
                    if impo1<>'':
                        impo1=round(impo1,2)
                    if impo2<>'':
                        impo2=round(impo2,2)
                    if impo3<>'':
                        impo3=round(impo3,2)
                    if impo4<>'':
                        impo4=round(impo4,2)
                    if impo5<>'':
                        impo5=round(impo5,2)

                    pp = ''
                    if move_line.partner_id.name:
                        pp = move_line.partner_id.name.encode('UTF-8','ignore')
                    nn = ''
                    if move_line.name:
                        nn = move_line.name.encode('UTF-8','ignore')
                    rr = ''
                    if move_line.move_id.name:
                        rr = move_line.move_id.name.encode('UTF-8','ignore')

                    righe.append([pp,nn,rr,move_line.date,scad1,impo1,scad2,impo2,scad3,impo3,scad4,impo4,scad5,impo5])
                    
        righe.append(['','','','','','','','','','','','','',''])
        righe.append(['TOTALE','','','','',round(tot_impo1,2),'',round(tot_impo2,2),'',round(tot_impo3,2),'',round(tot_impo4,2),'',round(tot_impo5,2)])
        line_total_altro=['ALTRO','',round(tot_impo1,2),round(tot_impo2,2),round(tot_impo3,2),round(tot_impo4,2),round(tot_impo5,2),'','','','','','','']
        righe.append(['','','','','','','','','','','','','',''])
        righe.append(['','','','','','','','','','','','','',''])
        tot_tot_impo1=tot_tot_impo1+round(tot_impo1,2)
        tot_tot_impo2=tot_tot_impo2+round(tot_impo2,2)
        tot_tot_impo3=tot_tot_impo3+round(tot_impo3,2)
        tot_tot_impo4=tot_tot_impo4+round(tot_impo4,2)
        tot_tot_impo5=tot_tot_impo5+round(tot_impo5,2)


        ###righe dei totali
        righe.append(['TOTALI','Saldo iniziale',array_inizio_str[0][3:5]+'/'+array_inizio_str[0][6:],array_inizio_str[1][3:5]+'/'+array_inizio_str[1][6:],array_inizio_str[2][3:5]+'/'+array_inizio_str[2][6:],array_inizio_str[3][3:5]+'/'+array_inizio_str[3][6:],array_inizio_str[4][3:5]+'/'+array_inizio_str[4][6:],'','','','','','',''])
        righe.append(line_total_fatt_v)
        righe.append(line_total_ord_v)
        righe.append(line_total_fatt_a)
        righe.append(line_total_ord_a)
        righe.append(line_total_altro)
        righe.append(['','','','','','','','','','','','','',''])
        righe.append(['TOTALE MESE',saldo_tot,round(tot_tot_impo1,2),round(tot_tot_impo2,2),round(tot_tot_impo3,2),round(tot_tot_impo4,2),round(tot_tot_impo5,2),'','','','','','',''])

        righe.append(['TOTALE CUMULATO',saldo_tot,round((saldo_tot+round(tot_tot_impo1,2)),2),round((saldo_tot+round(tot_tot_impo1,2)+round(tot_tot_impo2,2)),2),round((saldo_tot+round(tot_tot_impo1,2)+round(tot_tot_impo2,2)+round(tot_tot_impo3,2)),2),round((saldo_tot+round(tot_tot_impo1,2)+round(tot_tot_impo2,2)+round(tot_tot_impo3,2)+round(tot_tot_impo4,2)),2),round((saldo_tot+round(tot_tot_impo1,2)+round(tot_tot_impo2,2)+round(tot_tot_impo3,2)+round(tot_tot_impo4,2)+round(tot_tot_impo5,2)),2),'','','','','','',''])
        return righe

    @api.multi
    def export_report(self):
        self.ensure_one()

        base_style = xlwt.easyxf('align: wrap yes')
        workbook = xlwt.Workbook(encoding='utf8')
        for i, line in enumerate(self.prepare_report()):
            if line[0] in ('INCASSI-FATTURE','INCASSI-ORDINI','PAGAMENTI-FATTURE','PAGAMENTI-ORDINI','ALTRE SCADENZE','TOTALI'):
                worksheet = workbook.add_sheet(line[0])
                if line[0]=='TOTALI':
                    worksheet.set_horz_split_pos(1)
                else:
                    worksheet.set_horz_split_pos(2)
                linei = i
                worksheet.col(0).width = 20000
            rig = i-linei
            for j, cell in enumerate(line):
                if cell:
                    worksheet.write(rig, j, cell, base_style)
                    if isinstance(cell, basestring):
                        if 'Data ' in cell:
                            worksheet.col(j).width = 4000
                        elif 'Numero ' in cell or cell in ('Registrazione contabile', 'Nome'):
                            worksheet.col(j).width = 5500

        fp = StringIO()
        workbook.save(fp)
        fp.seek(0)
        self.write({'state': 'get', 'file': base64.encodestring(fp.getvalue())})
        fp.close()

        return {
            'name': "Cashflow Completato",
            'type': 'ir.actions.act_window',
            'res_model': 'previsione.in.out',
            'view_mode': 'form',
            'view_type': 'form',
            'res_id': self._ids[0],
            'views': [(False, 'form')],
            'target': 'new',
        }    

