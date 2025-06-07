# PowerPoint-raportin lataus
                st.markdown("---")
                st.subheader("📄 Luo mukautettu PowerPoint-raportti")
                
                # Valintamahdollisuudet
                st.markdown("**Valitse mitä osioita haluat mukaan raporttiin:**")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    include_summary = st.checkbox("🎯 Tuottavuustavoitteiden tulokset", value=True, help="Päivä- ja yötyöntekijöiden tavoitteiden täyttyminen")
                    include_monthly = st.checkbox("📊 Kuukauden yhteenveto", value=True, help="Päivittäisten tavoitteiden prosentit ja kuukauden ääripäät")
                    include_recommendations = st.checkbox("💡 Optimointisuositukset", value=True, help="Konkreettiset toimenpide-ehdotukset ja kokonaisarvio")
                
                with col2:
                    include_hourly = st.checkbox("⏰ Tuntikohtainen analyysi", value=False, help="Tuottavimmat ja vähiten tuottavat tunnit")
                    include_data_table = st.checkbox("📋 Päivittäisten tulosten taulukko", value=False, help="Yksityiskohtainen datataulukko kaikista päivistä")
                
                # Näytä kuinka monta diaa tulee
                selected_count = sum([include_summary, include_monthly, include_recommendations, include_hourly, include_data_table]) + 1  # +1 otsikkodia
                st.info(f"📄 Raportti sisältää {selected_count} diaa (sis. otsikkodia)")
                
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("🎯 Luo PowerPoint-raportti", type="primary", use_container_width=True):
                        # Tarkista että vähintään yksi osio on valittu
                        if not any([include_summary, include_monthly, include_recommendations, include_hourly, include_data_table]):
                            st.error("❌ Valitse vähintään yksi osio raporttiin!")
                        else:
                            with st.spinner("Luodaan PowerPoint-raporttia..."):
                                try:
                                    # Kerää valitut osiotimport streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import calendar
import io
import base64
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import plotly.io as pio

# Sivun konfiguraatio
st.set_page_config(
    page_title="Hälytysten Analyysihallinta",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Suomenkieliset nimet
FINNISH_MONTHS = [
    "Tammikuu", "Helmikuu", "Maaliskuu", "Huhtikuu", "Toukokuu", "Kesäkuu",
    "Heinäkuu", "Elokuu", "Syyskuu", "Lokakuu", "Marraskuu", "Joulukuu"
]

FINNISH_WEEKDAYS = ["Ma", "Ti", "Ke", "To", "Pe", "La", "Su"]
FINNISH_WEEKDAYS_LONG = ["Maanantai", "Tiistai", "Keskiviikko", "Torstai", "Perjantai", "Lauantai", "Sunnuntai"]

def get_worker_count(hour):
    """Laske työntekijämäärä tunnin perusteella"""
    workers = 0
    
    # Yövuoro 19:15-07:15 (2 henkilöä)
    if hour >= 19 or hour < 7:
        workers += 2
    
    # Aamuvuoro 07:00-17:00 (3 henkilöä)
    if hour >= 7 and hour < 17:
        workers += 3
    
    # Iltavuorot (portaittain)
    if hour >= 9 and hour < 19: workers += 1  # 09:15-19:15
    if hour >= 10 and hour < 20: workers += 1  # 10:00-20:00
    if hour >= 11 and hour < 21: workers += 1  # 11:00-21:00
    if hour >= 13 and hour < 23: workers += 1  # 13:00-23:00
    
    return workers

def get_finnish_weekday(date_obj):
    """Palauttaa suomenkielisen viikonpäivän nimen"""
    if pd.isna(date_obj):
        return "Tuntematon"
    weekday_num = date_obj.weekday()  # 0=Maanantai, 6=Sunnuntai
    return FINNISH_WEEKDAYS_LONG[weekday_num]

def get_finnish_month_name(month_num):
    """Palauttaa suomenkielisen kuukauden nimen"""
    return FINNISH_MONTHS[month_num - 1]

def create_powerpoint_report(daily_stats, hourly_stats, day_avg, night_avg, include_options):
    """Luo PowerPoint-raportti analyysin tuloksista valittujen osioiden mukaan"""
    try:
        # Luo uusi PowerPoint-esitys
        prs = Presentation()
        
        # 1. OTSIKKODIA (aina mukana)
        slide1 = prs.slides.add_slide(prs.slide_layouts[0])  # Title slide
        title = slide1.shapes.title
        subtitle = slide1.placeholders[1]
        
        title.text = "📊 Hälytysten Analyysiraportti"
        subtitle.text = f"Analysoitu: {datetime.now().strftime('%d.%m.%Y')}\n\nTuottavuustavoitteiden arviointi ja suositukset"
        
        # 2. YHTEENVETO-DIA (jos valittu)
        if include_options.get('summary', True):
            slide2 = prs.slides.add_slide(prs.slide_layouts[1])  # Title and content
            title2 = slide2.shapes.title
            title2.text = "🎯 Tuottavuustavoitteiden tulokset"
            
            # Lisää tekstikehys
            content = slide2.placeholders[1]
            tf = content.text_frame
            tf.clear()
            
            # Päivätyöntekijöiden tulos
            p1 = tf.paragraphs[0]
            p1.text = "🌅 PÄIVÄTYÖNTEKIJÄT (07-23)"
            p1.font.size = Pt(18)
            p1.font.bold = True
            
            p2 = tf.add_paragraph()
            day_status = "✅ TAVOITE SAAVUTETTU" if day_avg >= 5.1 else "❌ TAVOITE EI SAAVUTETTU"
            p2.text = f"Keskiarvo: {day_avg:.2f} inc/työnt./h\nTavoite: ≥5.1 inc/työnt./h\nTulos: {day_status}\nEro tavoitteeseen: {day_avg - 5.1:+.2f}"
            p2.font.size = Pt(14)
            
            # Yötyöntekijöiden tulos
            p3 = tf.add_paragraph()
            p3.text = "\n🌙 YÖTYÖNTEKIJÄT (23-07)"
            p3.font.size = Pt(18)
            p3.font.bold = True
            
            p4 = tf.add_paragraph()
            night_status = "✅ TAVOITE SAAVUTETTU" if night_avg >= 4.6 else "❌ TAVOITE EI SAAVUTETTU"
            p4.text = f"Keskiarvo: {night_avg:.2f} inc/työnt./h\nTavoite: ≥4.6 inc/työnt./h\nTulos: {night_status}\nEro tavoitteeseen: {night_avg - 4.6:+.2f}"
            p4.font.size = Pt(14)
        
        # 3. KUUKAUDEN TILASTOT (jos valittu)
        if include_options.get('monthly_stats', True):
            slide3 = prs.slides.add_slide(prs.slide_layouts[1])
            title3 = slide3.shapes.title
            title3.text = "📊 Kuukauden yhteenveto"
            
            content3 = slide3.placeholders[1]
            tf3 = content3.text_frame
            tf3.clear()
            
            # Laske tilastot
            day_target_met = len(daily_stats[daily_stats['day_target_met']])
            night_target_met = len(daily_stats[daily_stats['night_target_met']])
            total_days = len(daily_stats)
            
            max_day = daily_stats.loc[daily_stats['total_incidents'].idxmax()]
            min_day = daily_stats.loc[daily_stats['total_incidents'].idxmin()]
            
            p1 = tf3.paragraphs[0]
            p1.text = "📈 PÄIVITTÄISTEN TAVOITTEIDEN TÄYTTYMINEN"
            p1.font.size = Pt(16)
            p1.font.bold = True
            
            p2 = tf3.add_paragraph()
            p2.text = f"• Päivätyöntekijöiden tavoite täytetty: {day_target_met}/{total_days} päivää ({day_target_met/total_days*100:.1f}%)"
            p2.font.size = Pt(14)
            
            p3 = tf3.add_paragraph()
            p3.text = f"• Yötyöntekijöiden tavoite täytetty: {night_target_met}/{total_days} päivää ({night_target_met/total_days*100:.1f}%)"
            p3.font.size = Pt(14)
            
            p4 = tf3.add_paragraph()
            p4.text = f"\n📊 KUUKAUDEN ÄÄRIPÄÄT"
            p4.font.size = Pt(16)
            p4.font.bold = True
            
            p5 = tf3.add_paragraph()
            p5.text = f"• Kiireisin päivä: {max_day['day']:.0f}. ({max_day['day_name']}) - {max_day['total_incidents']:.0f} incidenttiä"
            p5.font.size = Pt(14)
            
            p6 = tf3.add_paragraph()
            p6.text = f"• Rauhallisin päivä: {min_day['day']:.0f}. ({min_day['day_name']}) - {min_day['total_incidents']:.0f} incidenttiä"
            p6.font.size = Pt(14)
        
        # 4. TUNTIKOHTAINEN ANALYYSI (jos valittu)
        if include_options.get('hourly_analysis', False):
            slide_hourly = prs.slides.add_slide(prs.slide_layouts[1])
            title_hourly = slide_hourly.shapes.title
            title_hourly.text = "⏰ Tuntikohtainen analyysi"
            
            content_hourly = slide_hourly.placeholders[1]
            tf_hourly = content_hourly.text_frame
            tf_hourly.clear()
            
            # Huipputunnit
            top_hours = hourly_stats.nlargest(3, 'incidents_per_worker')
            low_hours = hourly_stats.nsmallest(3, 'incidents_per_worker')
            
            p1 = tf_hourly.paragraphs[0]
            p1.text = "🔝 TUOTTAVIMMAT TUNNIT"
            p1.font.size = Pt(16)
            p1.font.bold = True
            
            for i, (_, row) in enumerate(top_hours.iterrows()):
                p = tf_hourly.add_paragraph()
                p.text = f"{i+1}. {row['hour_str']}: {row['incidents_per_worker']:.2f} inc/työnt./h ({row['worker_count']} työntekijää)"
                p.font.size = Pt(14)
            
            p_low = tf_hourly.add_paragraph()
            p_low.text = "\n📉 VÄHITEN TUOTTAVAT TUNNIT"
            p_low.font.size = Pt(16)
            p_low.font.bold = True
            
            for i, (_, row) in enumerate(low_hours.iterrows()):
                p = tf_hourly.add_paragraph()
                p.text = f"{i+1}. {row['hour_str']}: {row['incidents_per_worker']:.2f} inc/työnt./h ({row['worker_count']} työntekijää)"
                p.font.size = Pt(14)
        
        # 5. SUOSITUKSET (jos valittu)
        if include_options.get('recommendations', True):
            slide4 = prs.slides.add_slide(prs.slide_layouts[1])
            title4 = slide4.shapes.title
            title4.text = "💡 Optimointisuositukset"
            
            content4 = slide4.placeholders[1]
            tf4 = content4.text_frame
            tf4.clear()
            
            # Analysoi ongelmat
            day_problems = hourly_stats[
                (hourly_stats['hour'] >= 7) & 
                (hourly_stats['hour'] < 23) & 
                (hourly_stats['incidents_per_worker'] < 5.1)
            ]
            
            night_problems = hourly_stats[
                ((hourly_stats['hour'] >= 23) | (hourly_stats['hour'] < 7)) & 
                (hourly_stats['incidents_per_worker'] < 4.6)
            ]
            
            p1 = tf4.paragraphs[0]
            p1.text = "🎯 TOIMENPIDESUOSITUKSET"
            p1.font.size = Pt(18)
            p1.font.bold = True
            
            # Päivätyöntekijöiden suositukset
            p2 = tf4.add_paragraph()
            p2.text = "\n🌅 Päivätyöntekijät:"
            p2.font.size = Pt(16)
            p2.font.bold = True
            
            p3 = tf4.add_paragraph()
            if len(day_problems) > 0:
                problem_hours = ", ".join([f"{row['hour_str']}" for _, row in day_problems.iterrows()])
                p3.text = f"• Ongelmia {len(day_problems)} tunnissa: {problem_hours}\n• Suositus: Vähennä henkilöstöä ali-tuottavina aikoina tai siirrä tehtäviä"
            else:
                p3.text = "• ✅ Kaikki tunnit täyttävät tavoitteen!"
            p3.font.size = Pt(14)
            
            # Yötyöntekijöiden suositukset
            p4 = tf4.add_paragraph()
            p4.text = "\n🌙 Yötyöntekijät:"
            p4.font.size = Pt(16)
            p4.font.bold = True
            
            p5 = tf4.add_paragraph()
            if len(night_problems) > 0:
                problem_hours = ", ".join([f"{row['hour_str']}" for _, row in night_problems.iterrows()])
                p5.text = f"• Ongelmia {len(night_problems)} tunnissa: {problem_hours}\n• Suositus: Lisää henkilöstöä ongelmallisina aikoina"
            else:
                p5.text = "• ✅ Kaikki tunnit täyttävät tavoitteen!"
            p5.font.size = Pt(14)
            
            # Kokonaisarvio
            p6 = tf4.add_paragraph()
            p6.text = "\n📈 KOKONAISARVIO:"
            p6.font.size = Pt(16)
            p6.font.bold = True
            
            p7 = tf4.add_paragraph()
            if day_avg >= 5.1 and night_avg >= 4.6:
                overall_assessment = "🎉 Molemmat tuottavuustavoitteet saavutettu! Jatka samalla strategialla."
            elif day_avg >= 5.1:
                overall_assessment = "⚠️ Päivätyöntekijöiden tavoite saavutettu, mutta yötyöntekijät tarvitsevat parannusta."
            elif night_avg >= 4.6:
                overall_assessment = "⚠️ Yötyöntekijöiden tavoite saavutettu, mutta päivätyöntekijät tarvitsevat parannusta."
            else:
                overall_assessment = "❌ Kumpikaan tuottavuustavoite ei täyty. Tarvitaan merkittäviä toimenpiteitä."
            
            p7.text = overall_assessment
            p7.font.size = Pt(14)
        
        # 6. DATA-TAULUKKO (jos valittu)
        if include_options.get('data_table', False):
            slide5 = prs.slides.add_slide(prs.slide_layouts[1])
            title5 = slide5.shapes.title
            title5.text = "📋 Päivittäiset tulokset"
            
            # Lisää taulukko
            rows = min(len(daily_stats) + 1, 15)  # Max 14 datariviä + otsikko
            cols = 5
            
            left = Inches(0.5)
            top = Inches(1.5)
            width = Inches(9)
            height = Inches(5)
            
            table = slide5.shapes.add_table(rows, cols, left, top, width, height).table
            
            # Otsikot
            table.cell(0, 0).text = "Päivämäärä"
            table.cell(0, 1).text = "Viikonpäivä"
            table.cell(0, 2).text = "Incidentit"
            table.cell(0, 3).text = "Päivätyönt."
            table.cell(0, 4).text = "Yötyönt."
            
            # Data
            for i, (_, row) in enumerate(daily_stats.head(rows-1).iterrows()):
                table.cell(i+1, 0).text = str(row['date'])
                table.cell(i+1, 1).text = row['day_name']
                table.cell(i+1, 2).text = f"{row['total_incidents']:.0f}"
                
                day_status = "✅" if row['day_target_met'] else "❌"
                table.cell(i+1, 3).text = f"{row['day_shift_avg']:.2f} {day_status}"
                
                night_status = "✅" if row['night_target_met'] else "❌"
                table.cell(i+1, 4).text = f"{row['night_shift_avg']:.2f} {night_status}"
        
        # Tallenna muistiin
        pptx_buffer = io.BytesIO()
        prs.save(pptx_buffer)
        pptx_buffer.seek(0)
        
        return pptx_buffer
        
    except Exception as e:
        st.error(f"Virhe PowerPoint-raportin luonnissa: {str(e)}")
        return None

def create_calendar_view(daily_stats):
    """Luo kalenterinäkymä päivittäisistä tilastoista"""
    if len(daily_stats) == 0:
        return None
    
    # Muunna päivämäärät datetime-objekteiksi
    daily_stats = daily_stats.copy()
    daily_stats['date_obj'] = pd.to_datetime(daily_stats['date'])
    
    # Määritä kuukausi ja vuosi
    first_date = daily_stats['date_obj'].min()
    last_date = daily_stats['date_obj'].max()
    
    # Jos kaikki data on samalta kuukaudelta
    if first_date.month == last_date.month and first_date.year == last_date.year:
        month = first_date.month
        year = first_date.year
    else:
        # Käytä ensimmäistä kuukautta
        month = first_date.month
        year = first_date.year
    
    # Luo kuukauden kalenteri
    cal = calendar.monthcalendar(year, month)
    
    # Luo HTML-taulukko
    month_name = get_finnish_month_name(month)
    
    # Aloita kalenteri-HTML
    calendar_html = f"""
    <div style="margin: 20px 0; font-family: Arial, sans-serif;">
        <h3 style="text-align: center; margin-bottom: 20px; color: #1f77b4; font-size: 24px;">
            📅 {month_name} {year}
        </h3>
        <div style="text-align: center; margin-bottom: 15px; font-size: 14px;">
            <span style="color: #666;">Rauhallisin: </span>
            <span style="background-color: #d4edda; padding: 4px 12px; border-radius: 5px; font-weight: bold;">
                {daily_stats.loc[daily_stats['total_incidents'].idxmin(), 'day']:.0f}. päivä ({daily_stats['total_incidents'].min():.0f} inc)
            </span>
        </div>
        <table style="width: 100%; border-collapse: collapse; margin: 0 auto; max-width: 1000px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
            <thead>
                <tr style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;">
    """
    
    # Viikonpäivien otsikot
    for day_name in FINNISH_WEEKDAYS:
        calendar_html += f'<th style="padding: 15px 10px; text-align: center; font-weight: bold; font-size: 16px;">{day_name}</th>'
    
    calendar_html += """
                </tr>
            </thead>
            <tbody>
    """
    
    # Kalenterin rivit
    for week in cal:
        calendar_html += '<tr>'
        for day in week:
            if day == 0:
                # Tyhjä päivä
                calendar_html += '<td style="padding: 20px; border: 1px solid #e0e0e0; background-color: #f8f9fa; height: 85px;"></td>'
            else:
                # Etsi päivän data
                day_data = daily_stats[daily_stats['date_obj'].dt.day == day]
                
                if len(day_data) > 0:
                    row = day_data.iloc[0]
                    
                    # Määritä väri tavoitteiden perusteella
                    if row['day_target_met'] and row['night_target_met']:
                        bg_color = "#d4edda"  # Vihreä - molemmat tavoitteet täytetty
                        border_color = "#28a745"
                        border_width = "3px"
                    elif row['day_target_met'] or row['night_target_met']:
                        bg_color = "#fff3cd"  # Keltainen - yksi tavoite täytetty
                        border_color = "#ffc107"
                        border_width = "2px"
                    else:
                        bg_color = "#f8d7da"  # Punainen - kumpikaan tavoite ei täytetty
                        border_color = "#dc3545"
                        border_width = "2px"
                    
                    # Määritä P: ja Y: tekstien värit tavoitteiden mukaan
                    day_text_color = "#28a745" if row['day_target_met'] else "#dc3545"  # Vihreä jos tavoite täyttyy, muuten punainen
                    night_text_color = "#28a745" if row['night_target_met'] else "#dc3545"  # Vihreä jos tavoite täyttyy, muuten punainen
                    
                    calendar_html += f"""
                    <td style="padding: 10px; border: {border_width} solid {border_color}; background-color: {bg_color}; vertical-align: top; height: 85px; position: relative; transition: all 0.3s ease;">
                        <div style="font-weight: bold; font-size: 18px; margin-bottom: 6px; color: #333;">{day}</div>
                        <div style="font-size: 11px; line-height: 1.3;">
                            <div style="color: {day_text_color}; font-weight: bold; margin-bottom: 1px;">P: {row['day_shift_avg']:.2f}</div>
                            <div style="color: {night_text_color}; font-weight: bold; margin-bottom: 1px;">Y: {row['night_shift_avg']:.2f}</div>
                            <div style="color: #666; font-size: 10px; background-color: rgba(255,255,255,0.7); padding: 1px 3px; border-radius: 3px; display: inline-block;">{row['total_incidents']:.0f} inc</div>
                        </div>
                    </td>
                    """
                else:
                    # Ei dataa tälle päivälle
                    calendar_html += f"""
                    <td style="padding: 10px; border: 1px solid #dee2e6; background-color: #ffffff; vertical-align: top; height: 85px;">
                        <div style="font-weight: bold; color: #999; font-size: 16px; margin-bottom: 4px;">{day}</div>
                        <div style="font-size: 10px; color: #ccc; font-style: italic;">Ei dataa</div>
                    </td>
                    """
        calendar_html += '</tr>'
    
    calendar_html += """
            </tbody>
        </table>
        <div style="margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 8px;">
            <div style="text-align: center; font-size: 13px; color: #666; margin-bottom: 10px;">
                <strong>Selite:</strong>
            </div>
            <div style="display: flex; justify-content: center; gap: 25px; flex-wrap: wrap;">
                <span style="display: flex; align-items: center; gap: 8px;">
                    <span style="display: inline-block; width: 16px; height: 16px; background-color: #d4edda; border: 2px solid #28a745; border-radius: 3px;"></span>
                    <span style="font-size: 13px; font-weight: 500;">Molemmat tavoitteet täytetty</span>
                </span>
                <span style="display: flex; align-items: center; gap: 8px;">
                    <span style="display: inline-block; width: 16px; height: 16px; background-color: #fff3cd; border: 2px solid #ffc107; border-radius: 3px;"></span>
                    <span style="font-size: 13px; font-weight: 500;">Yksi tavoite täytetty</span>
                </span>
                <span style="display: flex; align-items: center; gap: 8px;">
                    <span style="display: inline-block; width: 16px; height: 16px; background-color: #f8d7da; border: 2px solid #dc3545; border-radius: 3px;"></span>
                    <span style="font-size: 13px; font-weight: 500;">Kumpikaan tavoite ei täytetty</span>
                </span>
                <span style="display: flex; align-items: center; gap: 8px;">
                    <span style="display: inline-block; width: 16px; height: 16px; background-color: #ffffff; border: 2px solid #dee2e6; border-radius: 3px;"></span>
                    <span style="font-size: 13px; font-weight: 500;">Ei dataa</span>
                </span>
            </div>
            <div style="text-align: center; margin-top: 12px; font-size: 12px; color: #888;">
                <strong style="color: #28a745;">P:</strong> Päivätyöntekijät (tavoite ≥5.1) | <strong style="color: #6f42c1;">Y:</strong> Yötyöntekijät (tavoite ≥4.6) | <strong>inc:</strong> Incidentit yhteensä
            </div>
        </div>
    </div>
    """
    
    return calendar_html

@st.cache_data
def process_data(df):
    """Käsittele Excel-data analyysiin"""
    try:
        # Tarkista että tarvittavat sarakkeet löytyvät
        required_columns = ['Hour', 'Incidents handled by agent']
        for col in required_columns:
            if col not in df.columns:
                st.error(f"Saraketta '{col}' ei löydy datasta. Tarkista Excel-tiedosto.")
                return None
        
        # Tee kopio datasta
        df_clean = df.copy()
        
        # Muunna Hour-sarake numeroiksi
        try:
            df_clean['Hour'] = pd.to_numeric(df_clean['Hour'], errors='coerce')
        except Exception as e:
            st.error(f"Virhe Hour-sarakkeen muunnossa: {str(e)}")
            return None
        
        # Muunna Incidents-sarake numeroiksi
        try:
            df_clean['Incidents handled by agent'] = pd.to_numeric(df_clean['Incidents handled by agent'], errors='coerce')
        except Exception as e:
            st.error(f"Virhe Incidents-sarakkeen muunnossa: {str(e)}")
            return None
        
        # Suodata vain validi data
        df_clean = df_clean[
            (df_clean['Hour'].notna()) & 
            (df_clean['Incidents handled by agent'].notna()) &
            (df_clean['Hour'] >= 0) & 
            (df_clean['Hour'] <= 23)
        ].copy()
        
        if len(df_clean) == 0:
            st.error("Ei validia dataa löydetty. Tarkista että Hour-sarake sisältää numeroita 0-23 ja Incidents-sarake sisältää numeroita.")
            return None
        
        # Lisää työntekijämäärät ja laskelmat
        df_clean['workers'] = df_clean['Hour'].apply(get_worker_count)
        df_clean['incidents_per_worker'] = df_clean['Incidents handled by agent'] / df_clean['workers']
        
        # Käsittele päivämäärät
        if 'Date' in df.columns:
            try:
                if df_clean['Date'].dtype in ['int64', 'float64']:
                    df_clean['date'] = pd.to_datetime('1900-01-01') + pd.to_timedelta(df_clean['Date'] - 2, unit='D')
                else:
                    df_clean['date'] = pd.to_datetime(df_clean['Date'], errors='coerce')
                
                if df_clean['date'].isna().all():
                    df_clean['date'] = datetime.now().date()
                    df_clean['date_str'] = df_clean['date'].astype(str)
                    df_clean['day_name'] = 'Tuntematon'
                    df_clean['day'] = 1
                else:
                    df_clean['date_str'] = df_clean['date'].dt.strftime('%Y-%m-%d')
                    df_clean['day_name'] = df_clean['date'].apply(get_finnish_weekday)
                    df_clean['day'] = df_clean['date'].dt.day
            except Exception as e:
                st.warning(f"Päivämäärien käsittely epäonnistui: {str(e)}. Käytetään oletuspäivämääriä.")
                df_clean['date'] = datetime.now().date()
                df_clean['date_str'] = df_clean['date'].astype(str)
                df_clean['day_name'] = 'Tuntematon'
                df_clean['day'] = 1
        else:
            df_clean['date'] = datetime.now().date()
            df_clean['date_str'] = df_clean['date'].astype(str)
            df_clean['day_name'] = 'Tuntematon'
            df_clean['day'] = 1
        
        return df_clean
        
    except Exception as e:
        st.error(f"Virhe datan käsittelyssä: {str(e)}")
        return None

def calculate_hourly_stats(df):
    """Laske tuntikohtaiset tilastot"""
    hourly_stats = []
    
    for hour in range(24):
        hour_data = df[df['Hour'] == hour]
        if len(hour_data) > 0:
            avg_incidents = hour_data['Incidents handled by agent'].mean()
            worker_count = get_worker_count(hour)
            avg_incidents_per_worker = avg_incidents / worker_count
            
            hourly_stats.append({
                'hour': hour,
                'hour_str': f"{hour:02d}:00",
                'avg_incidents': round(avg_incidents, 2),
                'worker_count': worker_count,
                'incidents_per_worker': round(avg_incidents_per_worker, 2),
                'days_count': len(hour_data)
            })
    
    return pd.DataFrame(hourly_stats)

def calculate_daily_stats(df):
    """Laske päivittäiset tilastot"""
    daily_stats = []
    
    for date_str in df['date_str'].unique():
        day_data = df[df['date_str'] == date_str]
        
        # Jaa päivä- ja yötyöntekijöihin
        day_shift = day_data[(day_data['Hour'] >= 7) & (day_data['Hour'] < 23)]
        night_shift = day_data[(day_data['Hour'] >= 23) | (day_data['Hour'] < 7)]
        
        day_shift_avg = day_shift['incidents_per_worker'].mean() if len(day_shift) > 0 else 0
        night_shift_avg = night_shift['incidents_per_worker'].mean() if len(night_shift) > 0 else 0
        
        daily_stats.append({
            'date': date_str,
            'day_name': day_data['day_name'].iloc[0] if len(day_data) > 0 else 'Tuntematon',
            'day': day_data['day'].iloc[0] if len(day_data) > 0 else 1,
            'total_incidents': day_data['Incidents handled by agent'].sum(),
            'day_shift_avg': round(day_shift_avg, 2),
            'night_shift_avg': round(night_shift_avg, 2),
            'day_target_met': day_shift_avg >= 5.1,
            'night_target_met': night_shift_avg >= 4.6
        })
    
    return pd.DataFrame(daily_stats)

def create_combined_chart(hourly_df):
    """Luo yhdistetty kaavio paremmilla tooltip-näkymillä"""
    fig = make_subplots(
        rows=1, cols=1,
        specs=[[{"secondary_y": True}]],
        subplot_titles=["Yhdistetty analyysi"]
    )
    
    # Pylväskaavio incidenteille - parannetulla tooltip
    fig.add_trace(
        go.Bar(
            x=hourly_df['hour_str'],
            y=hourly_df['avg_incidents'],
            name='Keskimääräiset incidentit',
            marker_color='lightblue',
            hovertemplate='<b>Kelloaika:</b> %{x}<br>' +
                         '<b>Keskimääräiset incidentit:</b> %{y:.2f}<br>' +
                         '<extra></extra>'
        ),
        secondary_y=False
    )
    
    # Viivakaavio incidenteille per työntekijä - parannetulla tooltip
    fig.add_trace(
        go.Scatter(
            x=hourly_df['hour_str'],
            y=hourly_df['incidents_per_worker'],
            name='Incidentit/työntekijä',
            line=dict(color='red', width=3),
            mode='lines+markers',
            marker=dict(size=8),
            hovertemplate='<b>Kelloaika:</b> %{x}<br>' +
                         '<b>Incidentit/työntekijä:</b> %{y:.2f}<br>' +
                         '<extra></extra>'
        ),
        secondary_y=True
    )
    
    # Viivakaavio työntekijämäärille - parannetulla tooltip
    fig.add_trace(
        go.Scatter(
            x=hourly_df['hour_str'],
            y=hourly_df['worker_count'],
            name='Työntekijämäärä',
            line=dict(color='green', width=2),
            mode='lines+markers',
            marker=dict(size=6),
            hovertemplate='<b>Kelloaika:</b> %{x}<br>' +
                         '<b>Työntekijämäärä:</b> %{y}<br>' +
                         '<extra></extra>'
        ),
        secondary_y=True
    )
    
    fig.update_xaxes(title_text="Kelloaika")
    fig.update_yaxes(title_text="Incidentit", secondary_y=False)
    fig.update_yaxes(title_text="Incidentit/työntekijä & Työntekijämäärä", secondary_y=True)
    
    fig.update_layout(
        height=500, 
        showlegend=True,
        hovermode='x unified',
        hoverlabel=dict(
            bgcolor="white",
            font_size=14,
            font_family="Arial"
        )
    )
    
    return fig

def main():
    # Otsikko
    st.title("📊 Hälytysten Analyysihallinta")
    st.markdown("**Lataa Excel-tiedosto ja saa automaattinen analyysi hälytysten määrästä suhteessa työntekijöihin**")
    
    # Sivupalkki
    with st.sidebar:
        st.header("⚙️ Asetukset")
        
        # Tiedoston lataus
        uploaded_file = st.file_uploader(
            "Lataa Excel-tiedosto",
            type=['xlsx', 'xls'],
            help="Tiedoston tulee sisältää sarakkeet: 'Hour', 'Incidents handled by agent', ja mahdollisesti 'Date'"
        )
        
        st.markdown("---")
        
        # PowerPoint-raportin lataus
        if 'processed_df' in locals() and processed_df is not None:
            st.subheader("📄 Lataa raportti")
            
            if st.button("🎯 Luo PowerPoint-raportti", type="primary"):
                with st.spinner("Luodaan PowerPoint-raporttia..."):
                    # Laske tarvittavat tilastot
                    daily_stats_for_ppt = calculate_daily_stats(processed_df)
                    hourly_stats_for_ppt = calculate_hourly_stats(processed_df)
                    
                    day_shift_data = processed_df[(processed_df['Hour'] >= 7) & (processed_df['Hour'] < 23)]
                    night_shift_data = processed_df[(processed_df['Hour'] >= 23) | (processed_df['Hour'] < 7)]
                    
                    day_avg_ppt = day_shift_data['incidents_per_worker'].mean() if len(day_shift_data) > 0 else 0
                    night_avg_ppt = night_shift_data['incidents_per_worker'].mean() if len(night_shift_data) > 0 else 0
                    
                    # Luo PowerPoint
                    pptx_buffer = create_powerpoint_report(daily_stats_for_ppt, hourly_stats_for_ppt, day_avg_ppt, night_avg_ppt)
                    
                    if pptx_buffer:
                        # Latausnappi
                        current_date = datetime.now().strftime("%Y%m%d")
                        filename = f"Hälytysten_Analyysi_{current_date}.pptx"
                        
                        st.download_button(
                            label="💾 Lataa PowerPoint-raportti",
                            data=pptx_buffer.getvalue(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="primary"
                        )
                        st.success("✅ PowerPoint-raportti valmis ladattavaksi!")
                    else:
                        st.error("❌ PowerPoint-raportin luonti epäonnistui.")
        
        st.markdown("---")
        
        # Vuorojen selitys
        st.subheader("🕐 Vuorojärjestely")
        st.markdown("""
        **Yövuoro** (19:15-07:15): 2 henkilöä  
        **Aamuvuoro** (07:00-17:00): 3 henkilöä  
        **Iltavuorot** (portaittain): 1-4 henkilöä
        - 09:15-19:15: +1 henkilö
        - 10:00-20:00: +1 henkilö  
        - 11:00-21:00: +1 henkilö
        - 13:00-23:00: +1 henkilö
        """)
        
        st.markdown("---")
        
        # Tuottavuustavoitteet
        st.subheader("🎯 Tuottavuustavoitteet")
        st.markdown("""
        **Päivätyöntekijät** (07-23): ≥5.1 inc/työnt./h  
        **Yötyöntekijät** (23-07): ≥4.6 inc/työnt./h
        """)
    
    # Pääsisältö
    if uploaded_file is not None:
        try:
            # Lue Excel-tiedosto
            df = pd.read_excel(uploaded_file)
            st.success(f"✅ Tiedosto ladattu! Löydettiin {len(df)} riviä dataa.")
            
            # Näytä datan otsikko
            with st.expander("📋 Näytä raakadata (ensimmäiset 10 riviä)"):
                st.dataframe(df.head(10))
                
                # Näytä sarakkeiden tietotyypit
                st.subheader("Sarakkeiden tietotyypit:")
                for col in df.columns:
                    st.write(f"- **{col}**: {df[col].dtype}")
            
            # Käsittele data
            processed_df = process_data(df)
            
            if processed_df is not None:
                st.success(f"✅ Data käsitelty onnistuneesti! {len(processed_df)} validia riviä.")
                
                # Laske tilastot
                hourly_stats = calculate_hourly_stats(processed_df)
                daily_stats = calculate_daily_stats(processed_df)
                
                # Tuottavuustavoitteiden analyysi
                day_shift_data = processed_df[(processed_df['Hour'] >= 7) & (processed_df['Hour'] < 23)]
                night_shift_data = processed_df[(processed_df['Hour'] >= 23) | (processed_df['Hour'] < 7)]
                
                day_avg = day_shift_data['incidents_per_worker'].mean() if len(day_shift_data) > 0 else 0
                night_avg = night_shift_data['incidents_per_worker'].mean() if len(night_shift_data) > 0 else 0
                
                # Tulosten näyttäminen
                st.header("🎯 Tuottavuustavoitteiden tulokset")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    day_status = "✅ SAAVUTETTU" if day_avg >= 5.1 else "❌ EI SAAVUTETTU"
                    day_color = "green" if day_avg >= 5.1 else "red"
                    st.markdown(f"""
                    <div style="padding: 20px; border: 2px solid {day_color}; border-radius: 10px; background-color: {'lightgreen' if day_avg >= 5.1 else 'lightcoral'};">
                        <h3>🌅 Päivätyöntekijät (07-23)</h3>
                        <p><strong>Keskiarvo:</strong> {day_avg:.2f} inc/työnt./h</p>
                        <p><strong>Tavoite:</strong> ≥5.1 inc/työnt./h</p>
                        <p><strong>Tulos:</strong> {day_status}</p>
                        <p><strong>Ero:</strong> {day_avg - 5.1:+.2f}</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    night_status = "✅ SAAVUTETTU" if night_avg >= 4.6 else "❌ EI SAAVUTETTU"
                    night_color = "green" if night_avg >= 4.6 else "red"
                    st.markdown(f"""
                    <div style="padding: 20px; border: 2px solid {night_color}; border-radius: 10px; background-color: {'lightgreen' if night_avg >= 4.6 else 'lightcoral'};">
                        <h3>🌙 Yötyöntekijät (23-07)</h3>
                        <p><strong>Keskiarvo:</strong> {night_avg:.2f} inc/työnt./h</p>
                        <p><strong>Tavoite:</strong> ≥4.6 inc/työnt./h</p>
                        <p><strong>Tulos:</strong> {night_status}</p>
                        <p><strong>Ero:</strong> {night_avg - 4.6:+.2f}</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Välilehdet eri näkymille
                tab1, tab2, tab3, tab4, tab5 = st.tabs([
                    "📊 Yhdistetty näkymä", 
                    "📈 Tuntikohtainen analyysi", 
                    "📅 Kuukausinäkymä",
                    "📋 Yksityiskohtaiset tilastot",
                    "💡 Suositukset"
                ])
                
                with tab1:
                    st.subheader("Yhdistetty analyysi")
                    if len(hourly_stats) > 0:
                        try:
                            fig_combined = create_combined_chart(hourly_stats)
                            st.plotly_chart(fig_combined, use_container_width=True)
                        except Exception as e:
                            st.error(f"Virhe kaavion luonnissa: {str(e)}")
                            st.info("Näytetään data taulukkona:")
                            st.dataframe(hourly_stats)
                    else:
                        st.warning("Ei dataa kaavion piirtämiseen.")
                
                with tab2:
                    st.subheader("Tuntikohtainen analyysi")
                    
                    if len(hourly_stats) > 0:
                        # Valitse näkymä
                        chart_type = st.selectbox(
                            "Valitse näkymä:",
                            ["Incidentit/työntekijä", "Kokonaisincidentit", "Työntekijämäärät"]
                        )
                        
                        try:
                            if chart_type == "Incidentit/työntekijä":
                                fig = px.line(
                                    hourly_stats, 
                                    x='hour_str', 
                                    y='incidents_per_worker',
                                    title='Incidentit per työntekijä tunnissa',
                                    markers=True,
                                    hover_data={
                                        'hour_str': False,
                                        'incidents_per_worker': ':.2f',
                                        'worker_count': True,
                                        'avg_incidents': ':.2f'
                                    }
                                )
                                fig.update_traces(
                                    hovertemplate='<b>Kelloaika:</b> %{x}<br>' +
                                                 '<b>Incidentit/työntekijä:</b> %{y:.2f}<br>' +
                                                 '<b>Työntekijämäärä:</b> %{customdata[0]}<br>' +
                                                 '<b>Keskimääräiset incidentit:</b> %{customdata[1]:.2f}<br>' +
                                                 '<extra></extra>',
                                    customdata=hourly_stats[['worker_count', 'avg_incidents']].values
                                )
                                fig.add_hline(y=5.1, line_dash="dash", line_color="red", 
                                             annotation_text="Päivätyöntekijöiden tavoite (5.1)")
                                fig.add_hline(y=4.6, line_dash="dash", line_color="blue", 
                                             annotation_text="Yötyöntekijöiden tavoite (4.6)")
                            
                            elif chart_type == "Kokonaisincidentit":
                                fig = px.bar(
                                    hourly_stats, 
                                    x='hour_str', 
                                    y='avg_incidents',
                                    title='Keskimääräiset incidentit tunneittain',
                                    hover_data={
                                        'hour_str': False,
                                        'avg_incidents': ':.2f',
                                        'worker_count': True,
                                        'incidents_per_worker': ':.2f'
                                    }
                                )
                                fig.update_traces(
                                    hovertemplate='<b>Kelloaika:</b> %{x}<br>' +
                                                 '<b>Keskimääräiset incidentit:</b> %{y:.2f}<br>' +
                                                 '<b>Työntekijämäärä:</b> %{customdata[0]}<br>' +
                                                 '<b>Incidentit/työntekijä:</b> %{customdata[1]:.2f}<br>' +
                                                 '<extra></extra>',
                                    customdata=hourly_stats[['worker_count', 'incidents_per_worker']].values
                                )
                            
                            else:  # Työntekijämäärät
                                fig = px.bar(
                                    hourly_stats, 
                                    x='hour_str', 
                                    y='worker_count',
                                    title='Työntekijämäärät tunneittain',
                                    hover_data={
                                        'hour_str': False,
                                        'worker_count': True,
                                        'avg_incidents': ':.2f',
                                        'incidents_per_worker': ':.2f'
                                    }
                                )
                                fig.update_traces(
                                    hovertemplate='<b>Kelloaika:</b> %{x}<br>' +
                                                 '<b>Työntekijämäärä:</b> %{y}<br>' +
                                                 '<b>Keskimääräiset incidentit:</b> %{customdata[0]:.2f}<br>' +
                                                 '<b>Incidentit/työntekijä:</b> %{customdata[1]:.2f}<br>' +
                                                 '<extra></extra>',
                                    customdata=hourly_stats[['avg_incidents', 'incidents_per_worker']].values
                                )
                            
                            # Yhteinen hover-tyyli kaikille kaavioille
                            fig.update_layout(
                                height=500,
                                hovermode='x unified',
                                hoverlabel=dict(
                                    bgcolor="white",
                                    font_size=14,
                                    font_family="Arial",
                                    bordercolor="gray"
                                )
                            )
                            st.plotly_chart(fig, use_container_width=True)
                            
                        except Exception as e:
                            st.error(f"Virhe kaavion luonnissa: {str(e)}")
                            st.info("Näytetään data taulukkona:")
                            st.dataframe(hourly_stats)
                    else:
                        st.warning("Ei dataa kaavion piirtämiseen.")
                
                with tab3:
                    st.subheader("📅 Kuukausinäkymä")
                    
                    if len(daily_stats) >= 1:
                        # Luo kalenterinäkymä
                        try:
                            calendar_html = create_calendar_view(daily_stats)
                            if calendar_html:
                                # Käytä korkeampaa height-arvoa jotta koko kalenteri mahtuu
                                import streamlit.components.v1 as components
                                components.html(calendar_html, height=900, scrolling=True)
                            else:
                                st.warning("Kalenterin luonti epäonnistui.")
                        except Exception as e:
                            st.error(f"Virhe kalenterin luonnissa: {str(e)}")
                            st.info("Näytetään data taulukkona:")
                            st.dataframe(daily_stats)
                        
                        # Kuukausistatistiikat
                        st.subheader("📊 Kuukauden yhteenveto")
                        col1, col2, col3, col4 = st.columns(4)
                        
                        day_target_met = len(daily_stats[daily_stats['day_target_met']]) 
                        night_target_met = len(daily_stats[daily_stats['night_target_met']])
                        total_days = len(daily_stats)
                        
                        with col1:
                            st.metric("Päivätyöntekijät", f"{day_target_met}/{total_days}", f"{day_target_met/total_days*100:.1f}%")
                        with col2:
                            st.metric("Yötyöntekijät", f"{night_target_met}/{total_days}", f"{night_target_met/total_days*100:.1f}%")
                        with col3:
                            max_day = daily_stats.loc[daily_stats['total_incidents'].idxmax()]
                            st.metric("Kiireisin päivä", f"{max_day['day']:.0f}. ({max_day['day_name']})", f"{max_day['total_incidents']:.0f} inc")
                        with col4:
                            min_day = daily_stats.loc[daily_stats['total_incidents'].idxmin()]
                            st.metric("Rauhallisin päivä", f"{min_day['day']:.0f}. ({min_day['day_name']})", f"{min_day['total_incidents']:.0f} inc")
                        
                        # Päivittäinen kehitys
                        try:
                            fig_daily = px.line(
                                daily_stats, 
                                x='date', 
                                y=['day_shift_avg', 'night_shift_avg'],
                                title='Päivittäinen kehitys',
                                labels={
                                    'value': 'Inc/työnt./h', 
                                    'variable': 'Vuoro',
                                    'date': 'Päivämäärä'
                                },
                                hover_data={
                                    'date': False,
                                    'value': ':.2f'
                                }
                            )
                            
                            # Muuta legendan nimet suomeksi ja paranna tooltip
                            fig_daily.for_each_trace(
                                lambda t: t.update(
                                    name='Päivätyöntekijät' if 'day_shift_avg' in t.name else 'Yötyöntekijät',
                                    hovertemplate='<b>Päivämäärä:</b> %{x}<br>' +
                                                 '<b>' + ('Päivätyöntekijät' if 'day_shift_avg' in t.name else 'Yötyöntekijät') + ':</b> %{y:.2f}<br>' +
                                                 '<extra></extra>'
                                )
                            )
                            
                            fig_daily.add_hline(y=5.1, line_dash="dash", line_color="red", 
                                              annotation_text="Päivätyöntekijöiden tavoite (5.1)")
                            fig_daily.add_hline(y=4.6, line_dash="dash", line_color="blue", 
                                              annotation_text="Yötyöntekijöiden tavoite (4.6)")
                            
                            fig_daily.update_layout(
                                hovermode='x unified',
                                hoverlabel=dict(
                                    bgcolor="white",
                                    font_size=14,
                                    font_family="Arial",
                                    bordercolor="gray"
                                )
                            )
                            st.plotly_chart(fig_daily, use_container_width=True)
                        except Exception as e:
                            st.error(f"Virhe päivittäisen kehityksen kaavion luonnissa: {str(e)}")
                            st.info("Näytetään data taulukkona:")
                            st.dataframe(daily_stats[['date', 'day_shift_avg', 'night_shift_avg']])
                        
                        # Päivittäinen taulukko
                        st.subheader("📋 Päivittäiset tulokset")
                        daily_display = daily_stats.copy()
                        daily_display['Päivätyöntekijät'] = daily_display.apply(
                            lambda x: f"{x['day_shift_avg']:.2f} {'✅' if x['day_target_met'] else '❌'}", axis=1
                        )
                        daily_display['Yötyöntekijät'] = daily_display.apply(
                            lambda x: f"{x['night_shift_avg']:.2f} {'✅' if x['night_target_met'] else '❌'}", axis=1
                        )
                        
                        st.dataframe(
                            daily_display[['date', 'day_name', 'total_incidents', 'Päivätyöntekijät', 'Yötyöntekijät']],
                            column_config={
                                'date': 'Päivämäärä',
                                'day_name': 'Viikonpäivä',
                                'total_incidents': 'Yhteensä inc.',
                                'Päivätyöntekijät': 'Päivätyöntekijät',
                                'Yötyöntekijät': 'Yötyöntekijät'
                            },
                            use_container_width=True
                        )
                    else:
                        st.info("Kuukausinäkymä vaatii vähintään yhden päivän dataa.")
                
                with tab4:
                    st.subheader("Tuntikohtaiset tilastot")
                    if len(hourly_stats) > 0:
                        st.dataframe(
                            hourly_stats,
                            column_config={
                                'hour_str': 'Kelloaika',
                                'avg_incidents': 'Keskim. incidentit',
                                'worker_count': 'Työntekijämäärä',
                                'incidents_per_worker': 'Inc/työnt./h',
                                'days_count': 'Päivien lukumäärä'
                            },
                            use_container_width=True
                        )
                    else:
                        st.warning("Ei tilastoja näytettäväksi.")
                
                with tab5:
                    st.subheader("💡 Optimointisuositukset")
                    
                    if len(hourly_stats) > 0:
                        # Ongelmatunnit päivätyöntekijöille
                        day_problems = hourly_stats[
                            (hourly_stats['hour'] >= 7) & 
                            (hourly_stats['hour'] < 23) & 
                            (hourly_stats['incidents_per_worker'] < 5.1)
                        ]
                        
                        # Ongelmatunnit yötyöntekijöille  
                        night_problems = hourly_stats[
                            ((hourly_stats['hour'] >= 23) | (hourly_stats['hour'] < 7)) & 
                            (hourly_stats['incidents_per_worker'] < 4.6)
                        ]
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.markdown("### 🌅 Päivätyöntekijät")
                            if len(day_problems) > 0:
                                st.error(f"Ongelmia {len(day_problems)} tunnissa:")
                                for _, row in day_problems.iterrows():
                                    st.write(f"- {row['hour_str']}: {row['incidents_per_worker']} inc/työnt./h")
                                st.markdown("**Suositus:** Vähennä henkilöstöä ali-tuottavina aikoina tai siirrä tehtäviä.")
                            else:
                                st.success("✅ Kaikki tunnit täyttävät tavoitteen!")
                        
                        with col2:
                            st.markdown("### 🌙 Yötyöntekijät")
                            if len(night_problems) > 0:
                                st.error(f"Ongelmia {len(night_problems)} tunnissa:")
                                for _, row in night_problems.iterrows():
                                    st.write(f"- {row['hour_str']}: {row['incidents_per_worker']} inc/työnt./h")
                                st.markdown("**Suositus:** Lisää henkilöstöä ongelmallisina aikoina.")
                            else:
                                st.success("✅ Kaikki tunnit täyttävät tavoitteen!")
                        
                        # Kokonaiskuva
                        st.markdown("### 📊 Kokonaisarvio")
                        if day_avg >= 5.1 and night_avg >= 4.6:
                            st.success("🎉 Molemmat tuottavuustavoitteet saavutettu! Jatka samalla strategialla.")
                        elif day_avg >= 5.1:
                            st.warning("⚠️ Päivätyöntekijöiden tavoite saavutettu, mutta yötyöntekijät tarvitsevat parannusta.")
                        elif night_avg >= 4.6:
                            st.warning("⚠️ Yötyöntekijöiden tavoite saavutettu, mutta päivätyöntekijät tarvitsevat parannusta.")
                        else:
                            st.error("❌ Kumpikaan tuottavuustavoite ei täyty. Tarvitaan merkittäviä toimenpiteitä.")
                    else:
                        st.warning("Ei dataa suositusten tekemiseen.")
                
                # PowerPoint-raportin lataus
                st.markdown("---")
                st.subheader("📄 Lataa PowerPoint-raportti")
                
                                    # Kerää valitut osiot
                                    include_options = {
                                        'summary': include_summary,
                                        'monthly_stats': include_monthly,
                                        'recommendations': include_recommendations,
                                        'hourly_analysis': include_hourly,
                                        'data_table': include_data_table
                                    }
                                    
                                    # Luo PowerPoint
                                    pptx_buffer = create_powerpoint_report(daily_stats, hourly_stats, day_avg, night_avg, include_options)
                                    
                                    if pptx_buffer:
                                        # Tallenna session stateen
                                        st.session_state['pptx_data'] = pptx_buffer.getvalue()
                                        st.session_state['pptx_slides'] = selected_count
                                        st.success(f"✅ PowerPoint-raportti luotu onnistuneesti! ({selected_count} diaa)")
                                    else:
                                        st.error("❌ PowerPoint-raportin luonti epäonnistui.")
                                except Exception as e:
                                    st.error(f"❌ Virhe: {str(e)}")
                
                # Latausnappi jos PowerPoint on luotu
                if 'pptx_data' in st.session_state:
                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        current_date = datetime.now().strftime("%Y%m%d")
                        slide_count = st.session_state.get('pptx_slides', 'X')
                        filename = f"Hälytysten_Analyysi_{current_date}_{slide_count}diaa.pptx"
                        
                        st.download_button(
                            label=f"💾 Lataa PowerPoint-raportti ({slide_count} diaa)",
                            data=st.session_state['pptx_data'],
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="secondary",
                            use_container_width=True
                        )
                        
                        st.info("💡 Vinkki: Klikkaa painiketta ladataksesi raportin tietokoneellesi")
                
                # Näytä esikatselu siitä mitä raportti sisältää
                if any([include_summary, include_monthly, include_recommendations, include_hourly, include_data_table]):
                    with st.expander("👀 Esikatselu: Mitä raportti sisältää"):
                        st.markdown("**📋 Raportin sisältö:**")
                        st.markdown("1. 📊 **Otsikkodia** (aina mukana)")
                        
                        slide_num = 2
                        if include_summary:
                            st.markdown(f"{slide_num}. 🎯 **Tuottavuustavoitteiden tulokset**")
                            st.markdown("   - Päivätyöntekijöiden keskiarvo ja tavoitteen täyttyminen")
                            st.markdown("   - Yötyöntekijöiden keskiarvo ja tavoitteen täyttyminen")
                            st.markdown("   - Selkeät ✅/❌ statukset ja erot tavoitteisiin")
                            slide_num += 1
                        
                        if include_monthly:
                            st.markdown(f"{slide_num}. 📊 **Kuukauden yhteenveto**")
                            st.markdown("   - Päivittäisten tavoitteiden täyttymisprosentit")
                            st.markdown("   - Kiireisin ja rauhallisin päivä")
                            st.markdown("   - Kuukauden tilastollinen yhteenveto")
                            slide_num += 1
                        
                        if include_hourly:
                            st.markdown(f"{slide_num}. ⏰ **Tuntikohtainen analyysi**")
                            st.markdown("   - Top 3 tuottavinta tuntia")
                            st.markdown("   - Top 3 vähiten tuottavaa tuntia")
                            st.markdown("   - Työntekijämäärät tunneittain")
                            slide_num += 1
                        
                        if include_recommendations:
                            st.markdown(f"{slide_num}. 💡 **Optimointisuositukset**")
                            st.markdown("   - Ongelmalliset tunnit vuoroittain")
                            st.markdown("   - Konkreettiset toimenpide-ehdotukset")
                            st.markdown("   - Kokonaisarvio ja strategiset suositukset")
                            slide_num += 1
                        
                        if include_data_table:
                            st.markdown(f"{slide_num}. 📋 **Päivittäisten tulosten taulukko**")
                            st.markdown("   - Jokaisen päivän yksityiskohtaiset tulokset")
                            st.markdown("   - ✅/❌ merkit tavoitteiden täyttymisestä")
                            st.markdown("   - Helppolukuinen taulukkomuoto")
                
                st.markdown("---")
        
        except Exception as e:
            st.error(f"Virhe tiedoston käsittelyssä: {str(e)}")
            st.info("Tarkista että Excel-tiedosto sisältää sarakkeet 'Hour' ja 'Incidents handled by agent' ja että ne sisältävät numeroita.")
    
    else:
        # Ohjeet kun ei tiedostoa ladattu
        st.info("👆 Lataa Excel-tiedosto sivupalkista aloittaaksesi analyysin.")
        
        st.markdown("---")
        st.subheader("📋 Käyttöohjeet")
        st.markdown("""
        1. **Lataa Excel-tiedosto** sivupalkista
        2. Tiedoston tulee sisältää vähintään sarakkeet:
           - `Hour` (0-23, numeroina)
           - `Incidents handled by agent` (määrä, numeroina)
           - `Date` (valinnainen, päivämäärille)
        3. **Tarkastele tuloksia** eri välilehdiltä:
           - 📊 Yhdistetty näkymä
           - 📈 Tuntikohtainen analyysi  
           - 📅 Kuukausinäkymä
           - 📋 Tilastot
           - 💡 Suositukset
        """)
        
        st.markdown("### 🎯 Mitä työkalu analysoi:")
        st.markdown("""
        - **Tuottavuustavoitteiden täyttyminen** vuoroittain
        - **Tuntikohtaiset kuormitukset** ja henkilöstötarpeet
        - **Päivittäiset suorituskykytrendit** 
        - **Optimointisuositukset** resurssien allokointiin
        - **Interaktiiviset visualisoinnit** helposti ymmärrettävässä muodossa
        """)

        # Näytä esimerkki oikeasta datamuodosta
        st.markdown("### 📝 Esimerkki oikeasta datamuodosta:")
        example_data = pd.DataFrame({
            'Date': ['2025-02-01', '2025-02-01', '2025-02-01'],
            'Hour': [0, 1, 2],
            'Incidents handled by agent': [9, 14, 16]
        })
        st.dataframe(example_data, use_container_width=True)

if __name__ == "__main__":
    main()
