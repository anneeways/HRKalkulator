import streamlit as st
import pandas as pd
import numpy as np
import json
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import os
import io

# Optional imports for exports - gracefully handle if not available
try:
    from groq import Groq
    GROQ_AVAILABLE = True
except ImportError:
    GROQ_AVAILABLE = False

try:
    from pptx import Presentation
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Configure Streamlit page
st.set_page_config(
    page_title="Comprehensive HR ROI Calculator",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 0.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        text-align: center;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
    }
</style>
""", unsafe_allow_html=True)

# Initialize Groq client (optional)
@st.cache_resource
def init_groq():
    if not GROQ_AVAILABLE:
        return None
    try:
        api_key = os.getenv("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY", "")
        if api_key:
            return Groq(api_key=api_key)
    except:
        pass
    return None

groq_client = init_groq()

# Industry Templates
INDUSTRY_TEMPLATES = {
    'technology': {
        'name': "Technology",
        'avg_salary': 110000,
        'current_turnover': 22,
        'replacement_cost': 2.0,
        'productivity_gain': 18,
        'retention_improvement': 30,
        'team_performance_gain': 15
    },
    'finance': {
        'name': "Financial Services",
        'avg_salary': 105000,
        'current_turnover': 15,
        'replacement_cost': 1.8,
        'productivity_gain': 12,
        'retention_improvement': 20,
        'team_performance_gain': 10
    },
    'healthcare': {
        'name': "Healthcare",
        'avg_salary': 85000,
        'current_turnover': 20,
        'replacement_cost': 1.6,
        'productivity_gain': 14,
        'retention_improvement': 25,
        'team_performance_gain': 12
    },
    'manufacturing': {
        'name': "Manufacturing",
        'avg_salary': 80000,
        'current_turnover': 16,
        'replacement_cost': 1.4,
        'productivity_gain': 16,
        'retention_improvement': 22,
        'team_performance_gain': 14
    },
    'consulting': {
        'name': "Consulting",
        'avg_salary': 120000,
        'current_turnover': 25,
        'replacement_cost': 2.2,
        'productivity_gain': 20,
        'retention_improvement': 35,
        'team_performance_gain': 18
    }
}

def format_currency(amount):
    """Format amount as currency"""
    return f"${amount:,.0f}"

def get_roi_color_status(roi):
    """Get color and status for ROI"""
    if roi >= 300:
        return "🟢 Excellent"
    elif roi >= 200:
        return "🟡 Good"
    elif roi >= 100:
        return "🟠 Moderate"
    else:
        return "🔴 Review Required"

def get_payback_color_status(months):
    """Get color and status for payback period"""
    if months <= 12:
        return "🟢 Fast"
    elif months <= 18:
        return "🟡 Moderate"
    elif months <= 24:
        return "🟠 Slow"
    else:
        return "🔴 Very Slow"

def calculate_roi(params):
    """Calculate Leadership ROI and related metrics"""
    # Program Costs
    participant_time_cost = (
        params['participants'] * 
        (params['avg_salary'] * 1.3 / 12) * 
        (params['time_commitment'] / 160) * 
        params['program_duration']
    )
    
    total_program_costs = (
        params['facilitator_costs'] + params['materials_costs'] + 
        params['venue_costs'] + params['travel_costs'] + 
        params['technology_costs'] + params['assessment_costs'] + 
        participant_time_cost
    )
    
    # Annual Benefits
    productivity_benefit = params['participants'] * params['avg_salary'] * (params['productivity_gain'] / 100)
    
    retention_savings = (
        params['participants'] * (params['current_turnover'] / 100) * 
        (params['retention_improvement'] / 100) * params['avg_salary'] * params['replacement_cost']
    )
    
    team_productivity_benefit = (
        params['participants'] * params['team_size'] * 
        (params['avg_salary'] * 0.7) * (params['team_performance_gain'] / 100)
    )
    
    promotion_benefit = (
        params['participants'] * 0.3 * (params['promotion_acceleration'] / 12) * 
        (params['avg_salary'] * 0.2)
    )
    
    decision_benefit = (
        params['participants'] * params['avg_salary'] * 0.1 * 
        (params['decision_quality_gain'] / 100)
    )
    
    total_annual_benefits = (
        productivity_benefit + retention_savings + team_productivity_benefit + 
        promotion_benefit + decision_benefit
    )
    
    # Multi-year analysis
    total_benefits = total_annual_benefits * params['analysis_years']
    net_benefit = total_benefits - total_program_costs
    roi = (net_benefit / total_program_costs) * 100 if total_program_costs > 0 else 0
    payback_months = (total_program_costs / (total_annual_benefits / 12)) if total_annual_benefits > 0 else float('inf')
    
    # NPV calculation (8% discount rate)
    discount_rate = 0.08
    npv = -total_program_costs
    for year in range(1, params['analysis_years'] + 1):
        npv += total_annual_benefits / ((1 + discount_rate) ** year)
    
    benefit_cost_ratio = total_benefits / total_program_costs if total_program_costs > 0 else 0
    
    return {
        'costs': {
            'facilitator': params['facilitator_costs'],
            'materials': params['materials_costs'],
            'venue': params['venue_costs'],
            'travel': params['travel_costs'],
            'technology': params['technology_costs'],
            'assessment': params['assessment_costs'],
            'participant_time': participant_time_cost,
            'total': total_program_costs
        },
        'benefits': {
            'productivity': productivity_benefit,
            'retention': retention_savings,
            'team_performance': team_productivity_benefit,
            'promotion': promotion_benefit,
            'decision_quality': decision_benefit,
            'total_annual': total_annual_benefits,
            'total_multi_year': total_benefits
        },
        'kpis': {
            'roi': roi,
            'payback_months': payback_months,
            'npv': npv,
            'net_benefit': net_benefit,
            'benefit_cost_ratio': benefit_cost_ratio
        }
    }

def calculate_recruiting_roi(params):
    """Calculate Recruiting ROI: Time to Hire × Cost per Hire optimization"""
    # Current state
    current_time_to_hire = params['current_time_to_hire']
    current_cost_per_hire = params['current_cost_per_hire']
    annual_hires = params['annual_hires']
    
    # Improvements
    time_reduction = params['time_to_hire_reduction']
    cost_reduction = params['cost_per_hire_reduction']
    quality_improvement = params['hire_quality_improvement']
    
    # Calculate savings
    improved_time_to_hire = current_time_to_hire * (1 - time_reduction / 100)
    improved_cost_per_hire = current_cost_per_hire * (1 - cost_reduction / 100)
    
    # Annual savings
    time_savings_cost = annual_hires * (current_time_to_hire - improved_time_to_hire) * params['daily_productivity_cost']
    direct_cost_savings = annual_hires * (current_cost_per_hire - improved_cost_per_hire)
    quality_value = annual_hires * params['avg_salary'] * (quality_improvement / 100) * 0.2
    
    total_annual_savings = time_savings_cost + direct_cost_savings + quality_value
    
    # Investment costs
    total_investment = params['recruiting_tech_investment'] + params['recruiter_training_cost'] + params['process_improvement_cost']
    
    roi = (total_annual_savings - total_investment) / total_investment * 100 if total_investment > 0 else 0
    
    return {
        'savings': {
            'time_savings': time_savings_cost,
            'cost_savings': direct_cost_savings,
            'quality_value': quality_value,
            'total_annual': total_annual_savings
        },
        'investment': total_investment,
        'roi': roi,
        'improved_metrics': {
            'time_to_hire': improved_time_to_hire,
            'cost_per_hire': improved_cost_per_hire
        }
    }

def calculate_onboarding_roi(params):
    """Calculate Onboarding ROI: Time to Productivity × Retention impact"""
    current_time_to_productivity = params['current_time_to_productivity']
    new_hire_retention_rate = params['new_hire_retention_rate']
    annual_new_hires = params['annual_new_hires']
    
    # Improvements
    productivity_acceleration = params['productivity_acceleration']
    retention_improvement = params['onboarding_retention_improvement']
    
    # Calculate benefits
    improved_time_to_productivity = current_time_to_productivity * (1 - productivity_acceleration / 100)
    productivity_months_saved = current_time_to_productivity - improved_time_to_productivity
    
    productivity_value_per_hire = productivity_months_saved * (params['avg_salary'] / 12) * 0.5
    total_productivity_value = annual_new_hires * productivity_value_per_hire
    
    improved_retention = new_hire_retention_rate + retention_improvement
    additional_retention = (improved_retention - new_hire_retention_rate) / 100
    retention_saves = annual_new_hires * additional_retention * params['avg_salary'] * params['replacement_cost']
    
    total_annual_benefits = total_productivity_value + retention_saves
    
    # Investment costs
    total_investment = params['onboarding_program_cost'] + params['mentor_training_cost'] + params['onboarding_tech_cost']
    
    roi = (total_annual_benefits - total_investment) / total_investment * 100 if total_investment > 0 else 0
    
    return {
        'benefits': {
            'productivity_value': total_productivity_value,
            'retention_value': retention_saves,
            'total_annual': total_annual_benefits
        },
        'investment': total_investment,
        'roi': roi,
        'improved_metrics': {
            'time_to_productivity': improved_time_to_productivity,
            'retention_rate': improved_retention
        }
    }

def calculate_retention_roi(params):
    """Calculate Retention ROI: Engagement Score × Turnover cost reduction"""
    current_engagement_score = params['current_engagement_score']
    current_turnover_rate = params['current_turnover']
    total_employees = params['total_employees']
    
    # Improvements
    engagement_improvement = params['engagement_improvement']
    turnover_reduction = params['retention_improvement']
    
    # Calculate benefits
    improved_engagement = min(10, current_engagement_score + engagement_improvement)
    engagement_productivity_boost = engagement_improvement * 0.03
    
    productivity_benefit = total_employees * params['avg_salary'] * engagement_productivity_boost
    
    current_annual_turnover = total_employees * (current_turnover_rate / 100)
    reduced_turnover = current_annual_turnover * (turnover_reduction / 100)
    turnover_cost_savings = reduced_turnover * params['avg_salary'] * params['replacement_cost']
    
    absenteeism_reduction = total_employees * params['avg_salary'] * 0.02 * engagement_improvement
    customer_satisfaction_boost = total_employees * 1000 * engagement_improvement
    
    total_annual_benefits = productivity_benefit + turnover_cost_savings + absenteeism_reduction + customer_satisfaction_boost
    
    # Investment costs
    total_investment = params['engagement_program_cost'] + params['survey_analytics_cost'] + params['manager_training_cost']
    
    roi = (total_annual_benefits - total_investment) / total_investment * 100 if total_investment > 0 else 0
    
    return {
        'benefits': {
            'productivity_boost': productivity_benefit,
            'turnover_savings': turnover_cost_savings,
            'absenteeism_reduction': absenteeism_reduction,
            'customer_value': customer_satisfaction_boost,
            'total_annual': total_annual_benefits
        },
        'investment': total_investment,
        'roi': roi,
        'improved_metrics': {
            'engagement_score': improved_engagement,
            'turnover_rate': current_turnover_rate * (1 - turnover_reduction / 100)
        }
    }

def calculate_development_roi(params):
    """Calculate Development ROI: Internal Mobility × Succession planning"""
    internal_mobility_rate = params['internal_mobility_rate']
    succession_readiness = params['succession_readiness']
    total_leadership_positions = params['total_leadership_positions']
    
    # Improvements
    mobility_improvement = params['mobility_improvement']
    succession_improvement = params['succession_improvement']
    
    # Calculate benefits
    improved_mobility_rate = internal_mobility_rate + mobility_improvement
    annual_leadership_openings = total_leadership_positions * 0.15
    
    external_hire_cost = params['external_leadership_hire_cost']
    internal_promotion_cost = params['internal_promotion_cost']
    cost_per_internal_hire = external_hire_cost - internal_promotion_cost
    
    additional_internal_hires = annual_leadership_openings * (mobility_improvement / 100)
    hiring_cost_savings = additional_internal_hires * cost_per_internal_hire
    
    improved_succession_readiness = succession_readiness + succession_improvement
    reduced_succession_risk = (succession_improvement / 100) * total_leadership_positions
    succession_risk_value = reduced_succession_risk * params['avg_salary'] * 2
    
    development_participants = params['development_participants']
    performance_improvement = development_participants * params['avg_salary'] * (params['development_performance_gain'] / 100)
    
    development_retention_boost = development_participants * (params['development_retention_boost'] / 100)
    retention_value = development_retention_boost * params['avg_salary'] * params['replacement_cost']
    
    total_annual_benefits = hiring_cost_savings + succession_risk_value + performance_improvement + retention_value
    
    # Investment costs
    total_investment = params['development_program_cost'] + params['mentoring_program_cost'] + params['succession_planning_cost']
    
    roi = (total_annual_benefits - total_investment) / total_investment * 100 if total_investment > 0 else 0
    
    return {
        'benefits': {
            'hiring_savings': hiring_cost_savings,
            'succession_value': succession_risk_value,
            'performance_gains': performance_improvement,
            'retention_value': retention_value,
            'total_annual': total_annual_benefits
        },
        'investment': total_investment,
        'roi': roi,
        'improved_metrics': {
            'mobility_rate': improved_mobility_rate,
            'succession_readiness': improved_succession_readiness
        }
    }

def calculate_knowledge_transfer_roi(params):
    """Calculate Knowledge Transfer ROI: Alumni value × Knowledge preservation"""
    retiring_employees_annual = params['retiring_employees_annual']
    average_tenure_retirees = params['average_tenure_retirees']
    knowledge_loss_percentage = params['knowledge_loss_percentage']
    
    # Improvements
    knowledge_capture_rate = params['knowledge_capture_improvement']
    alumni_network_value = params['alumni_network_value']
    
    # Calculate knowledge value
    knowledge_per_employee = params['avg_salary'] * average_tenure_retirees * 0.1
    knowledge_preserved = retiring_employees_annual * knowledge_per_employee * (knowledge_capture_rate / 100)
    
    # Alumni network benefits
    active_alumni_connections = params['active_alumni_connections']
    annual_alumni_value = active_alumni_connections * alumni_network_value
    
    # Reduced onboarding time for replacements
    replacement_onboarding_savings = retiring_employees_annual * params['knowledge_transfer_time_savings'] * (params['avg_salary'] / 12)
    
    # Innovation and best practices sharing
    innovation_value = params['innovation_from_knowledge_transfer']
    
    total_annual_benefits = knowledge_preserved + annual_alumni_value + replacement_onboarding_savings + innovation_value
    
    # Investment costs
    total_investment = params['knowledge_capture_system_cost'] + params['documentation_program_cost'] + params['alumni_network_cost']
    
    roi = (total_annual_benefits - total_investment) / total_investment * 100 if total_investment > 0 else 0
    
    return {
        'benefits': {
            'knowledge_preserved': knowledge_preserved,
            'alumni_value': annual_alumni_value,
            'onboarding_savings': replacement_onboarding_savings,
            'innovation_value': innovation_value,
            'total_annual': total_annual_benefits
        },
        'investment': total_investment,
        'roi': roi,
        'knowledge_metrics': {
            'knowledge_loss_prevented': knowledge_capture_rate,
            'alumni_connections': active_alumni_connections
        }
    }

def get_ai_insights(comprehensive_results, params):
    """Get AI-powered insights using Groq"""
    if not groq_client:
        return "AI insights unavailable (Groq API key not configured)"
    
    models_to_try = ["llama3-70b-8192", "llama3-8b-8192", "gemma-7b-it"]
    
    overall_roi = comprehensive_results["overall"]["total_roi"]
    prompt = f"""
    Analyze this comprehensive HR ROI strategy:
    
    OVERALL METRICS:
    - Total ROI: {overall_roi:.1f}%
    - Total Investment: {format_currency(comprehensive_results["overall"]["total_investment"])}
    - Total Annual Benefits: {format_currency(comprehensive_results["overall"]["total_annual_benefits"])}
    
    MODULE PERFORMANCE:
    - Leadership ROI: {comprehensive_results['leadership']['kpis']['roi']:.1f}%
    - Recruiting ROI: {comprehensive_results['recruiting']['roi']:.1f}%
    - Onboarding ROI: {comprehensive_results['onboarding']['roi']:.1f}%
    - Retention ROI: {comprehensive_results['retention']['roi']:.1f}%
    - Development ROI: {comprehensive_results['development']['roi']:.1f}%
    - Knowledge Transfer ROI: {comprehensive_results['knowledge_transfer']['roi']:.1f}%
    
    Provide strategic recommendations for implementation priority and synergies.
    """
    
    for model in models_to_try:
        try:
            response = groq_client.chat.completions.create(
                messages=[{"role": "user", "content": prompt}],
                model=model,
                max_tokens=600
            )
            return response.choices[0].message.content
        except Exception as e:
            if "model_decommissioned" in str(e) or "not found" in str(e):
                continue
            else:
                return f"AI insights unavailable: {str(e)}"
    
    return "AI insights unavailable: All models are currently unavailable"

def create_text_report(comprehensive_results, params):
    """Create a text-based business case report"""
    overall_roi = comprehensive_results["overall"]["total_roi"]
    
    report = f"""
COMPREHENSIVE HR ROI BUSINESS CASE
Generated: {datetime.now().strftime('%B %d, %Y')}

EXECUTIVE SUMMARY
================
This comprehensive business case analyzes the return on investment across six critical HR modules: 
Leadership Development, Recruiting Optimization, Onboarding Excellence, Employee Retention, 
Talent Development, and Knowledge Transfer.

Overall Findings:
• Total Investment: {format_currency(comprehensive_results["overall"]["total_investment"])}
• Total Annual Benefits: {format_currency(comprehensive_results["overall"]["total_annual_benefits"])}
• Overall ROI: {overall_roi:.0f}%
• Net Annual Benefit: {format_currency(comprehensive_results["overall"]["net_annual_benefit"])}

MODULE PERFORMANCE SUMMARY
==========================
Leadership Development: {comprehensive_results['leadership']['kpis']['roi']:.0f}% ROI
Investment: {format_currency(comprehensive_results['leadership']['costs']['total'])}
Annual Benefits: {format_currency(comprehensive_results['leadership']['benefits']['total_annual'])}

Recruiting Optimization: {comprehensive_results['recruiting']['roi']:.0f}% ROI
Investment: {format_currency(comprehensive_results['recruiting']['investment'])}
Annual Benefits: {format_currency(comprehensive_results['recruiting']['savings']['total_annual'])}

Onboarding Excellence: {comprehensive_results['onboarding']['roi']:.0f}% ROI
Investment: {format_currency(comprehensive_results['onboarding']['investment'])}
Annual Benefits: {format_currency(comprehensive_results['onboarding']['benefits']['total_annual'])}

Employee Retention: {comprehensive_results['retention']['roi']:.0f}% ROI
Investment: {format_currency(comprehensive_results['retention']['investment'])}
Annual Benefits: {format_currency(comprehensive_results['retention']['benefits']['total_annual'])}

Talent Development: {comprehensive_results['development']['roi']:.0f}% ROI
Investment: {format_currency(comprehensive_results['development']['investment'])}
Annual Benefits: {format_currency(comprehensive_results['development']['benefits']['total_annual'])}

Knowledge Transfer: {comprehensive_results['knowledge_transfer']['roi']:.0f}% ROI
Investment: {format_currency(comprehensive_results['knowledge_transfer']['investment'])}
Annual Benefits: {format_currency(comprehensive_results['knowledge_transfer']['benefits']['total_annual'])}

STRATEGIC RECOMMENDATIONS
=========================
Business Case Assessment: {"STRONG COMPREHENSIVE HR STRATEGY - Proceed with phased implementation" if overall_roi >= 200 else "MODERATE HR STRATEGY - Focus on high-ROI modules first" if overall_roi >= 100 else "REVIEW HR STRATEGY - Optimize assumptions and module selection"}

Implementation Priority:
• Phase 1: Launch highest ROI modules (>200% ROI) immediately
• Phase 2: Implement moderate ROI modules (100-200% ROI) within 6 months
• Phase 3: Optimize and scale successful programs

Key Success Factors:
• Establish robust measurement and analytics framework
• Secure executive sponsorship and change management support
• Create cross-module synergies and integration points
• Implement pilot programs before full-scale rollout
• Monitor and adjust programs based on real-world results

Generated by Comprehensive HR ROI Calculator
"""
    return report

def main():
    # Header
    st.markdown("""
    <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                padding: 2rem; border-radius: 10px; margin-bottom: 2rem;'>
        <h1 style='color: white; margin: 0; font-size: 2.5rem;'>🎯 Comprehensive HR ROI Calculator</h1>
        <p style='color: rgba(255,255,255,0.8); margin: 0.5rem 0 0 0; font-size: 1.2rem;'>
            Calculate ROI across Leadership, Recruiting, Onboarding, Retention, Development & Knowledge Transfer
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Show availability status
    col1, col2, col3 = st.columns(3)
    with col1:
        if GROQ_AVAILABLE and groq_client:
            st.success("✅ AI Insights Available")
        else:
            st.warning("⚠️ AI Insights: Add Groq API key")
    
    with col2:
        if PPTX_AVAILABLE:
            st.success("✅ PowerPoint Export Available")
        else:
            st.info("📊 PowerPoint: Use text export instead")
    
    with col3:
        if REPORTLAB_AVAILABLE:
            st.success("✅ PDF Export Available")
        else:
            st.info("📄 PDF: Use text export instead")
    
    # Initialize session state
    if 'params' not in st.session_state:
        st.session_state.params = {
            # Leadership Program Parameters
            'participants': 20,
            'program_duration': 6,
            'avg_salary': 95000,
            'time_commitment': 15,
            'analysis_years': 3,
            'facilitator_costs': 75000,
            'materials_costs': 15000,
            'venue_costs': 25000,
            'travel_costs': 30000,
            'technology_costs': 12000,
            'assessment_costs': 8000,
            'productivity_gain': 15,
            'retention_improvement': 25,
            'promotion_acceleration': 6,
            'team_performance_gain': 12,
            'decision_quality_gain': 20,
            'current_turnover': 18,
            'replacement_cost': 1.5,
            'team_size': 8,
            
            # Recruiting ROI Parameters
            'current_time_to_hire': 45,
            'current_cost_per_hire': 5000,
            'annual_hires': 50,
            'time_to_hire_reduction': 30,
            'cost_per_hire_reduction': 20,
            'hire_quality_improvement': 15,
            'daily_productivity_cost': 400,
            'recruiting_tech_investment': 25000,
            'recruiter_training_cost': 15000,
            'process_improvement_cost': 10000,
            
            # Onboarding ROI Parameters
            'current_time_to_productivity': 3,
            'new_hire_retention_rate': 75,
            'annual_new_hires': 50,
            'productivity_acceleration': 40,
            'onboarding_retention_improvement': 15,
            'onboarding_program_cost': 30000,
            'mentor_training_cost': 20000,
            'onboarding_tech_cost': 15000,
            
            # Retention ROI Parameters
            'current_engagement_score': 6.5,
            'total_employees': 500,
            'engagement_improvement': 1.5,
            'engagement_program_cost': 50000,
            'survey_analytics_cost': 15000,
            'manager_training_cost': 25000,
            
            # Development ROI Parameters
            'internal_mobility_rate': 60,
            'succession_readiness': 40,
            'total_leadership_positions': 50,
            'mobility_improvement': 20,
            'succession_improvement': 30,
            'external_leadership_hire_cost': 75000,
            'internal_promotion_cost': 15000,
            'development_participants': 100,
            'development_performance_gain': 12,
            'development_retention_boost': 25,
            'development_program_cost': 80000,
            'mentoring_program_cost': 30000,
            'succession_planning_cost': 20000,
            
            # Knowledge Transfer ROI Parameters
            'retiring_employees_annual': 25,
            'average_tenure_retirees': 15,
            'knowledge_loss_percentage': 70,
            'knowledge_capture_improvement': 50,
            'alumni_network_value': 2000,
            'active_alumni_connections': 100,
            'knowledge_transfer_time_savings': 2,
            'innovation_from_knowledge_transfer': 50000,
            'knowledge_capture_system_cost': 40000,
            'documentation_program_cost': 25000,
            'alumni_network_cost': 15000
        }
    
    # Sidebar for inputs
    with st.sidebar:
        st.header("📊 Program Configuration")
        
        # Industry Templates
        st.subheader("🏭 Industry Templates")
        industry_options = [''] + [template['name'] for template in INDUSTRY_TEMPLATES.values()]
        selected_industry = st.selectbox("Select Industry Template", industry_options)
        
        if selected_industry:
            template_key = None
            for key, template in INDUSTRY_TEMPLATES.items():
                if template['name'] == selected_industry:
                    template_key = key
                    break
            
            if template_key and st.button("Apply Template"):
                template = INDUSTRY_TEMPLATES[template_key]
                st.session_state.params.update({
                    'avg_salary': template['avg_salary'],
                    'current_turnover': template['current_turnover'],
                    'replacement_cost': template['replacement_cost'],
                    'productivity_gain': template['productivity_gain'],
                    'retention_improvement': template['retention_improvement'],
                    'team_performance_gain': template['team_performance_gain']
                })
                st.success(f"Applied {selected_industry} template!")
        
        st.divider()
        
        # Program Parameters
        st.subheader("📈 Program Parameters")
        st.session_state.params['participants'] = st.number_input(
            "Participants", min_value=1, value=st.session_state.params['participants']
        )
        st.session_state.params['avg_salary'] = st.number_input(
            "Average Salary ($)", min_value=0, value=st.session_state.params['avg_salary'], step=5000
        )
        st.session_state.params['total_employees'] = st.number_input(
            "Total Employees", min_value=1, value=st.session_state.params['total_employees']
        )
    
    # Calculate all results
    leadership_results = calculate_roi(st.session_state.params)
    recruiting_results = calculate_recruiting_roi(st.session_state.params)
    onboarding_results = calculate_onboarding_roi(st.session_state.params)
    retention_results = calculate_retention_roi(st.session_state.params)
    development_results = calculate_development_roi(st.session_state.params)
    knowledge_results = calculate_knowledge_transfer_roi(st.session_state.params)
    
    # Overall ROI Summary
    total_annual_benefits = (
        leadership_results['benefits']['total_annual'] +
        recruiting_results['savings']['total_annual'] +
        onboarding_results['benefits']['total_annual'] +
        retention_results['benefits']['total_annual'] +
        development_results['benefits']['total_annual'] +
        knowledge_results['benefits']['total_annual']
    )
    
    total_investment = (
        leadership_results['costs']['total'] +
        recruiting_results['investment'] +
        onboarding_results['investment'] +
        retention_results['investment'] +
        development_results['investment'] +
        knowledge_results['investment']
    )
    
    overall_roi = ((total_annual_benefits - total_investment) / total_investment * 100) if total_investment > 0 else 0
    
    # Overall Summary Dashboard
    st.subheader("🎯 Overall HR ROI Dashboard")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "Total ROI",
            f"{overall_roi:.0f}%",
            delta="🟢 Excellent" if overall_roi >= 200 else "🟡 Good" if overall_roi >= 100 else "🔴 Review"
        )
    
    with col2:
        st.metric(
            "Total Investment",
            format_currency(total_investment)
        )
    
    with col3:
        st.metric(
            "Annual Benefits",
            format_currency(total_annual_benefits)
        )
    
    with col4:
        net_benefit = total_annual_benefits - total_investment
        st.metric(
            "Net Annual Benefit",
            format_currency(net_benefit)
        )
    
    # ROI Comparison Chart
    roi_data = {
        'Module': ['Leadership', 'Recruiting', 'Onboarding', 'Retention', 'Development', 'Knowledge Transfer'],
        'ROI (%)': [
            leadership_results['kpis']['roi'],
            recruiting_results['roi'],
            onboarding_results['roi'],
            retention_results['roi'],
            development_results['roi'],
            knowledge_results['roi']
        ]
    }
    
    fig_roi = px.bar(
        x=roi_data['Module'],
        y=roi_data['ROI (%)'],
        title="ROI Comparison Across HR Modules",
        color=roi_data['ROI (%)'],
        color_continuous_scale="Viridis"
    )
    fig_roi.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig_roi, use_container_width=True)
    
    st.divider()
    
    # Main content tabs
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "🎯 Leadership", "🔍 Recruiting", "🚀 Onboarding", 
        "💝 Retention", "📈 Development", "🧠 Knowledge", 
        "⚙️ Configuration", "🤖 AI Insights"
    ])
    
    with tab1:
        st.subheader("🎯 Leadership Development ROI")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Leadership ROI", f"{leadership_results['kpis']['roi']:.0f}%", 
                     delta=get_roi_color_status(leadership_results['kpis']['roi']))
        
        with col2:
            st.metric("Payback Period", f"{leadership_results['kpis']['payback_months']:.1f} months",
                     delta=get_payback_color_status(leadership_results['kpis']['payback_months']))
        
        with col3:
            st.metric("Investment", format_currency(leadership_results['costs']['total']))
        
        with col4:
            st.metric("Annual Benefits", format_currency(leadership_results['benefits']['total_annual']))
        
        # Charts
        col1, col2 = st.columns(2)
        
        with col1:
            cost_data = {
                'Category': ['Facilitator', 'Materials', 'Venue', 'Travel', 'Technology', 'Assessment', 'Time'],
                'Amount': [
                    leadership_results['costs']['facilitator'],
                    leadership_results['costs']['materials'],
                    leadership_results['costs']['venue'],
                    leadership_results['costs']['travel'],
                    leadership_results['costs']['technology'],
                    leadership_results['costs']['assessment'],
                    leadership_results['costs']['participant_time']
                ]
            }
            
            fig_costs = px.pie(values=cost_data['Amount'], names=cost_data['Category'], 
                              title="Investment Breakdown")
            st.plotly_chart(fig_costs, use_container_width=True)
        
        with col2:
            benefit_data = {
                'Category': ['Productivity', 'Retention', 'Team Performance', 'Promotions', 'Decision Quality'],
                'Amount': [
                    leadership_results['benefits']['productivity'],
                    leadership_results['benefits']['retention'],
                    leadership_results['benefits']['team_performance'],
                    leadership_results['benefits']['promotion'],
                    leadership_results['benefits']['decision_quality']
                ]
            }
            
            fig_benefits = px.bar(x=benefit_data['Category'], y=benefit_data['Amount'], 
                                 title="Annual Benefits")
            fig_benefits.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_benefits, use_container_width=True)
    
    with tab2:
        st.subheader("🔍 Recruiting ROI Analysis")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Recruiting ROI", f"{recruiting_results['roi']:.0f}%")
        
        with col2:
            st.metric("Time to Hire", f"{recruiting_results['improved_metrics']['time_to_hire']:.0f} days",
                     delta=f"-{st.session_state.params['current_time_to_hire'] - recruiting_results['improved_metrics']['time_to_hire']:.0f}")
        
        with col3:
            st.metric("Cost per Hire", format_currency(recruiting_results['improved_metrics']['cost_per_hire']))
        
        with col4:
            st.metric("Annual Savings", format_currency(recruiting_results['savings']['total_annual']))
    
    with tab3:
        st.subheader("🚀 Onboarding ROI Analysis")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Onboarding ROI", f"{onboarding_results['roi']:.0f}%")
        
        with col2:
            st.metric("Time to Productivity", f"{onboarding_results['improved_metrics']['time_to_productivity']:.1f} months")
        
        with col3:
            st.metric("Retention Rate", f"{onboarding_results['improved_metrics']['retention_rate']:.0f}%")
        
        with col4:
            st.metric("Annual Benefits", format_currency(onboarding_results['benefits']['total_annual']))
    
    with tab4:
        st.subheader("💝 Employee Retention ROI")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Retention ROI", f"{retention_results['roi']:.0f}%")
        
        with col2:
            st.metric("Engagement Score", f"{retention_results['improved_metrics']['engagement_score']:.1f}/10")
        
        with col3:
            st.metric("Turnover Rate", f"{retention_results['improved_metrics']['turnover_rate']:.1f}%")
        
        with col4:
            st.metric("Annual Benefits", format_currency(retention_results['benefits']['total_annual']))
    
    with tab5:
        st.subheader("📈 Employee Development ROI")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Development ROI", f"{development_results['roi']:.0f}%")
        
        with col2:
            st.metric("Internal Mobility", f"{development_results['improved_metrics']['mobility_rate']:.0f}%")
        
        with col3:
            st.metric("Succession Readiness", f"{development_results['improved_metrics']['succession_readiness']:.0f}%")
        
        with col4:
            st.metric("Annual Benefits", format_currency(development_results['benefits']['total_annual']))
    
    with tab6:
        st.subheader("🧠 Knowledge Transfer ROI")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Knowledge ROI", f"{knowledge_results['roi']:.0f}%")
        
        with col2:
            st.metric("Knowledge Preserved", f"{knowledge_results['knowledge_metrics']['knowledge_loss_prevented']:.0f}%")
        
        with col3:
            st.metric("Alumni Connections", f"{knowledge_results['knowledge_metrics']['alumni_connections']}")
        
        with col4:
            st.metric("Annual Benefits", format_currency(knowledge_results['benefits']['total_annual']))
    
    with tab7:
        st.subheader("⚙️ Configuration")
        
        # Quick configuration options
        st.write("**Quick Adjustments:**")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.session_state.params['productivity_gain'] = st.slider(
                "Productivity Improvement (%)", 0, 50, st.session_state.params['productivity_gain']
            )
            st.session_state.params['retention_improvement'] = st.slider(
                "Retention Improvement (%)", 0, 50, st.session_state.params['retention_improvement']
            )
        
        with col2:
            st.session_state.params['current_turnover'] = st.slider(
                "Current Turnover Rate (%)", 0, 50, st.session_state.params['current_turnover']
            )
            st.session_state.params['replacement_cost'] = st.slider(
                "Replacement Cost Multiple", 0.5, 3.0, st.session_state.params['replacement_cost'], 0.1
            )
    
    with tab8:
        st.subheader("🤖 AI-Powered Insights")
        
        comprehensive_results = {
            'overall': {
                'total_roi': overall_roi,
                'total_investment': total_investment,
                'total_annual_benefits': total_annual_benefits,
                'net_annual_benefit': total_annual_benefits - total_investment
            },
            'leadership': leadership_results,
            'recruiting': recruiting_results,
            'onboarding': onboarding_results,
            'retention': retention_results,
            'development': development_results,
            'knowledge_transfer': knowledge_results
        }
        
        if GROQ_AVAILABLE and groq_client:
            if st.button("Generate AI Analysis", type="primary"):
                with st.spinner("Analyzing your HR ROI strategy..."):
                    insights = get_ai_insights(comprehensive_results, st.session_state.params)
                    st.markdown(insights)
        else:
            st.info("AI insights require Groq API key. Add your key in Streamlit Secrets to enable this feature.")
        
        st.divider()
        
        # Module comparison table
        comparison_data = pd.DataFrame({
            'Module': ['Leadership', 'Recruiting', 'Onboarding', 'Retention', 'Development', 'Knowledge'],
            'ROI (%)': [
                leadership_results['kpis']['roi'],
                recruiting_results['roi'],
                onboarding_results['roi'],
                retention_results['roi'],
                development_results['roi'],
                knowledge_results['roi']
            ],
            'Investment': [
                leadership_results['costs']['total'],
                recruiting_results['investment'],
                onboarding_results['investment'],
                retention_results['investment'],
                development_results['investment'],
                knowledge_results['investment']
            ]
        })
        
        comparison_data = comparison_data.sort_values('ROI (%)', ascending=False)
        st.subheader("🏆 Module Ranking by ROI")
        st.dataframe(comparison_data, use_container_width=True)
    
    # Export functionality
    st.divider()
    st.subheader("📄 Export Business Case")
    
    comprehensive_results = {
        'overall': {
            'total_roi': overall_roi,
            'total_investment': total_investment,
            'total_annual_benefits': total_annual_benefits,
            'net_annual_benefit': total_annual_benefits - total_investment
        },
        'leadership': leadership_results,
        'recruiting': recruiting_results,
        'onboarding': onboarding_results,
        'retention': retention_results,
        'development': development_results,
        'knowledge_transfer': knowledge_results
    }
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📋 Export Text Report", type="primary"):
            text_report = create_text_report(comprehensive_results, st.session_state.params)
            st.download_button(
                label="📥 Download Text Report",
                data=text_report,
                file_name=f"hr_roi_business_case_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                mime="text/plain"
            )
    
    with col2:
        if st.button("📊 Export Data (JSON)", type="secondary"):
            export_data = {
                'parameters': st.session_state.params,
                'results': comprehensive_results,
                'timestamp': datetime.now().isoformat()
            }
            st.download_button(
                label="📥 Download JSON",
                data=json.dumps(export_data, indent=2, default=str),
                file_name=f"hr_roi_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json"
            )
    
    with col3:
        recommendation = (
            "STRONG HR STRATEGY - Proceed with implementation" if overall_roi >= 200 else
            "MODERATE HR STRATEGY - Focus on high-ROI modules" if overall_roi >= 100 else
            "REVIEW HR STRATEGY - Optimize assumptions"
        )
        st.success(f"**Recommendation:** {recommendation}")

if __name__ == "__main__":
    main()
