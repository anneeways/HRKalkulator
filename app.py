import streamlit as st
import pandas as pd
import numpy as np
import json
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# Configure Streamlit page
st.set_page_config(
    page_title="HR ROI Calculator",
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
    .initiative-card {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #007bff;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Initiative Templates - Focus on individual HR programs
INITIATIVE_TEMPLATES = {
    'leadership_development': {
        'name': "Leadership Development Program",
        'description': "Comprehensive leadership training with coaching and assessments",
        'participants': 25,
        'program_duration': 6,
        'avg_salary': 100000,
        'facilitator_costs': 80000,
        'materials_costs': 20000,
        'venue_costs': 30000,
        'productivity_gain': 18,
        'retention_improvement': 30,
        'team_performance_gain': 15,
        'typical_roi': "250-400%"
    },
    'executive_coaching': {
        'name': "Executive Coaching Initiative", 
        'description': "1-on-1 coaching for senior leaders",
        'participants': 10,
        'program_duration': 12,
        'avg_salary': 150000,
        'facilitator_costs': 120000,
        'materials_costs': 5000,
        'venue_costs': 0,
        'productivity_gain': 25,
        'retention_improvement': 40,
        'team_performance_gain': 20,
        'typical_roi': "300-500%"
    },
    'recruiting_optimization': {
        'name': "Recruiting Process Optimization",
        'description': "Technology and process improvements for faster, better hiring",
        'annual_hires': 50,
        'current_time_to_hire': 45,
        'current_cost_per_hire': 5000,
        'time_to_hire_reduction': 35,
        'cost_per_hire_reduction': 25,
        'hire_quality_improvement': 20,
        'recruiting_tech_investment': 30000,
        'typical_roi': "200-350%"
    },
    'onboarding_excellence': {
        'name': "Structured Onboarding Program",
        'description': "Comprehensive new hire integration and training",
        'annual_new_hires': 50,
        'current_time_to_productivity': 4,
        'productivity_acceleration': 50,
        'new_hire_retention_rate': 75,
        'onboarding_retention_improvement': 20,
        'onboarding_program_cost': 40000,
        'typical_roi': "300-450%"
    },
    'engagement_retention': {
        'name': "Employee Engagement & Retention",
        'description': "Surveys, feedback systems, and engagement initiatives",
        'total_employees': 500,
        'current_engagement_score': 6.5,
        'current_turnover': 18,
        'engagement_improvement': 2.0,
        'retention_improvement': 30,
        'engagement_program_cost': 60000,
        'typical_roi': "250-400%"
    },
    'talent_development': {
        'name': "Internal Talent Development",
        'description': "Skills training, mentoring, and succession planning",
        'development_participants': 100,
        'internal_mobility_rate': 60,
        'succession_readiness': 40,
        'mobility_improvement': 25,
        'succession_improvement': 35,
        'development_program_cost': 100000,
        'typical_roi': "200-300%"
    }
}

def format_currency(amount):
    """Format amount as currency"""
    return f"${amount:,.0f}"

def get_roi_status(roi):
    """Get status and color for ROI"""
    if roi >= 300:
        return "🟢 Excellent (300%+)"
    elif roi >= 200:
        return "🟡 Good (200-299%)"
    elif roi >= 100:
        return "🟠 Moderate (100-199%)"
    else:
        return "🔴 Needs Review (<100%)"

def calculate_leadership_roi(params):
    """Calculate Leadership Development ROI"""
    # Costs
    participant_time_cost = (
        params['participants'] * 
        (params['avg_salary'] * 1.3 / 12) * 
        (params.get('time_commitment', 20) / 160) * 
        params['program_duration']
    )
    
    total_costs = (
        params['facilitator_costs'] + 
        params['materials_costs'] + 
        params['venue_costs'] + 
        params.get('travel_costs', 20000) +
        participant_time_cost
    )
    
    # Annual Benefits
    productivity_benefit = params['participants'] * params['avg_salary'] * (params['productivity_gain'] / 100)
    
    current_turnover = params.get('current_turnover', 18)
    replacement_cost = params.get('replacement_cost', 1.5)
    retention_savings = (
        params['participants'] * (current_turnover / 100) * 
        (params['retention_improvement'] / 100) * 
        params['avg_salary'] * replacement_cost
    )
    
    team_size = params.get('team_size', 8)
    team_benefit = (
        params['participants'] * team_size * 
        (params['avg_salary'] * 0.7) * (params['team_performance_gain'] / 100)
    )
    
    total_annual_benefits = productivity_benefit + retention_savings + team_benefit
    
    # Calculate ROI
    roi = ((total_annual_benefits - total_costs) / total_costs * 100) if total_costs > 0 else 0
    payback_months = (total_costs / (total_annual_benefits / 12)) if total_annual_benefits > 0 else 0
    
    return {
        'total_costs': total_costs,
        'annual_benefits': total_annual_benefits,
        'roi': roi,
        'payback_months': payback_months,
        'benefit_breakdown': {
            'productivity': productivity_benefit,
            'retention': retention_savings,
            'team_performance': team_benefit
        }
    }

def calculate_recruiting_roi(params):
    """Calculate Recruiting ROI"""
    # Current metrics
    current_time = params['current_time_to_hire']
    current_cost = params['current_cost_per_hire']
    annual_hires = params['annual_hires']
    
    # Improvements
    time_reduction = params['time_to_hire_reduction'] / 100
    cost_reduction = params['cost_per_hire_reduction'] / 100
    quality_improvement = params['hire_quality_improvement'] / 100
    
    # Calculate savings
    time_savings = annual_hires * current_time * time_reduction * params.get('daily_productivity_cost', 400)
    cost_savings = annual_hires * current_cost * cost_reduction
    quality_value = annual_hires * params.get('avg_salary', 95000) * quality_improvement * 0.15
    
    total_annual_savings = time_savings + cost_savings + quality_value
    total_investment = params['recruiting_tech_investment'] + params.get('training_costs', 15000)
    
    roi = ((total_annual_savings - total_investment) / total_investment * 100) if total_investment > 0 else 0
    
    return {
        'total_investment': total_investment,
        'annual_savings': total_annual_savings,
        'roi': roi,
        'improved_metrics': {
            'new_time_to_hire': current_time * (1 - time_reduction),
            'new_cost_per_hire': current_cost * (1 - cost_reduction)
        },
        'savings_breakdown': {
            'time_savings': time_savings,
            'cost_savings': cost_savings,
            'quality_value': quality_value
        }
    }

def calculate_onboarding_roi(params):
    """Calculate Onboarding ROI"""
    annual_hires = params['annual_new_hires']
    current_productivity_time = params['current_time_to_productivity']
    acceleration = params['productivity_acceleration'] / 100
    
    # Productivity improvement
    months_saved = current_productivity_time * acceleration
    productivity_value = annual_hires * months_saved * (params.get('avg_salary', 95000) / 12) * 0.6
    
    # Retention improvement
    retention_improvement = params['onboarding_retention_improvement'] / 100
    retention_value = annual_hires * retention_improvement * params.get('avg_salary', 95000) * 1.5
    
    total_annual_benefits = productivity_value + retention_value
    total_investment = params['onboarding_program_cost'] + params.get('training_costs', 15000)
    
    roi = ((total_annual_benefits - total_investment) / total_investment * 100) if total_investment > 0 else 0
    
    return {
        'total_investment': total_investment,
        'annual_benefits': total_annual_benefits,
        'roi': roi,
        'benefit_breakdown': {
            'productivity_value': productivity_value,
            'retention_value': retention_value
        }
    }

def calculate_engagement_roi(params):
    """Calculate Engagement & Retention ROI"""
    total_employees = params['total_employees']
    current_turnover = params['current_turnover'] / 100
    engagement_improvement = params['engagement_improvement']
    retention_improvement = params['retention_improvement'] / 100
    
    # Benefits
    productivity_boost = total_employees * params.get('avg_salary', 95000) * (engagement_improvement * 0.02)
    turnover_reduction = total_employees * current_turnover * retention_improvement
    turnover_savings = turnover_reduction * params.get('avg_salary', 95000) * 1.5
    
    total_annual_benefits = productivity_boost + turnover_savings
    total_investment = params['engagement_program_cost'] + params.get('survey_costs', 15000)
    
    roi = ((total_annual_benefits - total_investment) / total_investment * 100) if total_investment > 0 else 0
    
    return {
        'total_investment': total_investment,
        'annual_benefits': total_annual_benefits,
        'roi': roi,
        'benefit_breakdown': {
            'productivity_boost': productivity_boost,
            'turnover_savings': turnover_savings
        }
    }

def calculate_development_roi(params):
    """Calculate Talent Development ROI"""
    participants = params['development_participants']
    mobility_improvement = params['mobility_improvement'] / 100
    
    # Benefits
    performance_gains = participants * params.get('avg_salary', 95000) * 0.12
    internal_hiring_savings = participants * mobility_improvement * 25000  # vs external hiring
    retention_boost = participants * 0.2 * params.get('avg_salary', 95000) * 1.5
    
    total_annual_benefits = performance_gains + internal_hiring_savings + retention_boost
    total_investment = params['development_program_cost']
    
    roi = ((total_annual_benefits - total_investment) / total_investment * 100) if total_investment > 0 else 0
    
    return {
        'total_investment': total_investment,
        'annual_benefits': total_annual_benefits,
        'roi': roi,
        'benefit_breakdown': {
            'performance_gains': performance_gains,
            'hiring_savings': internal_hiring_savings,
            'retention_boost': retention_boost
        }
    }

def main():
    # Header
    st.markdown("""
    <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                padding: 2rem; border-radius: 10px; margin-bottom: 2rem;'>
        <h1 style='color: white; margin: 0; font-size: 2.5rem;'>🎯 HR ROI Calculator</h1>
        <p style='color: rgba(255,255,255,0.8); margin: 0.5rem 0 0 0; font-size: 1.2rem;'>
            Calculate ROI for Individual HR Initiatives
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state
    if 'selected_initiatives' not in st.session_state:
        st.session_state.selected_initiatives = []
    if 'params' not in st.session_state:
        st.session_state.params = {}
    
    # Sidebar for initiative selection
    with st.sidebar:
        st.header("🎯 Select HR Initiatives")
        
        # Initiative Templates
        st.subheader("📋 Available Templates")
        
        for key, template in INITIATIVE_TEMPLATES.items():
            with st.expander(f"📊 {template['name']}"):
                st.write(f"**Description:** {template['description']}")
                st.write(f"**Typical ROI:** {template['typical_roi']}")
                
                if st.button(f"Add {template['name']}", key=f"add_{key}"):
                    if key not in st.session_state.selected_initiatives:
                        st.session_state.selected_initiatives.append(key)
                        st.session_state.params[key] = template.copy()
                        st.success(f"Added {template['name']}!")
                        st.rerun()
        
        st.divider()
        
        # Currently selected initiatives
        st.subheader("✅ Selected Initiatives")
        if st.session_state.selected_initiatives:
            for initiative in st.session_state.selected_initiatives:
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"• {INITIATIVE_TEMPLATES[initiative]['name']}")
                with col2:
                    if st.button("❌", key=f"remove_{initiative}"):
                        st.session_state.selected_initiatives.remove(initiative)
                        if initiative in st.session_state.params:
                            del st.session_state.params[initiative]
                        st.rerun()
        else:
            st.info("No initiatives selected yet")
    
    # Main content
    if not st.session_state.selected_initiatives:
        st.info("👈 Please select one or more HR initiatives from the sidebar to begin calculating ROI.")
        return
    
    # Create tabs for each selected initiative
    if len(st.session_state.selected_initiatives) == 1:
        # Single initiative - no tabs needed
        initiative_key = st.session_state.selected_initiatives[0]
        display_initiative(initiative_key)
    else:
        # Multiple initiatives - create tabs
        tab_names = [INITIATIVE_TEMPLATES[key]['name'] for key in st.session_state.selected_initiatives]
        tabs = st.tabs(tab_names)
        
        for i, initiative_key in enumerate(st.session_state.selected_initiatives):
            with tabs[i]:
                display_initiative(initiative_key)
    
    # Overall summary if multiple initiatives
    if len(st.session_state.selected_initiatives) > 1:
        st.divider()
        display_overall_summary()

def display_initiative(initiative_key):
    """Display interface for a specific initiative"""
    template = INITIATIVE_TEMPLATES[initiative_key]
    params = st.session_state.params[initiative_key]
    
    st.subheader(f"📊 {template['name']}")
    st.write(template['description'])
    
    # Parameters input section
    with st.expander("⚙️ Adjust Parameters", expanded=True):
        col1, col2 = st.columns(2)
        
        # Update parameters based on initiative type
        if initiative_key == 'leadership_development' or initiative_key == 'executive_coaching':
            with col1:
                params['participants'] = st.number_input(
                    "Number of Participants", 
                    min_value=1, 
                    value=params['participants'],
                    key=f"participants_{initiative_key}"
                )
                params['avg_salary'] = st.number_input(
                    "Average Salary ($)", 
                    min_value=0, 
                    value=params['avg_salary'], 
                    step=5000,
                    key=f"salary_{initiative_key}"
                )
                params['program_duration'] = st.number_input(
                    "Program Duration (months)", 
                    min_value=1, 
                    value=params['program_duration'],
                    key=f"duration_{initiative_key}"
                )
            
            with col2:
                params['productivity_gain'] = st.slider(
                    "Productivity Improvement (%)", 
                    0, 50, 
                    params['productivity_gain'],
                    key=f"productivity_{initiative_key}"
                )
                params['retention_improvement'] = st.slider(
                    "Retention Improvement (%)", 
                    0, 50, 
                    params['retention_improvement'],
                    key=f"retention_{initiative_key}"
                )
                params['team_performance_gain'] = st.slider(
                    "Team Performance Gain (%)", 
                    0, 30, 
                    params['team_performance_gain'],
                    key=f"team_{initiative_key}"
                )
            
            # Calculate and display results
            results = calculate_leadership_roi(params)
            
        elif initiative_key == 'recruiting_optimization':
            with col1:
                params['annual_hires'] = st.number_input(
                    "Annual Hires", 
                    min_value=1, 
                    value=params['annual_hires'],
                    key=f"hires_{initiative_key}"
                )
                params['current_time_to_hire'] = st.number_input(
                    "Current Time to Hire (days)", 
                    min_value=1, 
                    value=params['current_time_to_hire'],
                    key=f"time_{initiative_key}"
                )
                params['current_cost_per_hire'] = st.number_input(
                    "Current Cost per Hire ($)", 
                    min_value=0, 
                    value=params['current_cost_per_hire'],
                    key=f"cost_{initiative_key}"
                )
            
            with col2:
                params['time_to_hire_reduction'] = st.slider(
                    "Time Reduction (%)", 
                    0, 50, 
                    params['time_to_hire_reduction'],
                    key=f"time_reduction_{initiative_key}"
                )
                params['cost_per_hire_reduction'] = st.slider(
                    "Cost Reduction (%)", 
                    0, 50, 
                    params['cost_per_hire_reduction'],
                    key=f"cost_reduction_{initiative_key}"
                )
                params['hire_quality_improvement'] = st.slider(
                    "Quality Improvement (%)", 
                    0, 30, 
                    params['hire_quality_improvement'],
                    key=f"quality_{initiative_key}"
                )
            
            results = calculate_recruiting_roi(params)
            
        elif initiative_key == 'onboarding_excellence':
            with col1:
                params['annual_new_hires'] = st.number_input(
                    "Annual New Hires", 
                    min_value=1, 
                    value=params['annual_new_hires'],
                    key=f"new_hires_{initiative_key}"
                )
                params['current_time_to_productivity'] = st.number_input(
                    "Current Time to Productivity (months)", 
                    min_value=1, 
                    value=params['current_time_to_productivity'],
                    key=f"productivity_time_{initiative_key}"
                )
            
            with col2:
                params['productivity_acceleration'] = st.slider(
                    "Productivity Acceleration (%)", 
                    0, 70, 
                    params['productivity_acceleration'],
                    key=f"acceleration_{initiative_key}"
                )
                params['onboarding_retention_improvement'] = st.slider(
                    "Retention Improvement (%)", 
                    0, 40, 
                    params['onboarding_retention_improvement'],
                    key=f"onboard_retention_{initiative_key}"
                )
            
            results = calculate_onboarding_roi(params)
            
        elif initiative_key == 'engagement_retention':
            with col1:
                params['total_employees'] = st.number_input(
                    "Total Employees", 
                    min_value=1, 
                    value=params['total_employees'],
                    key=f"employees_{initiative_key}"
                )
                params['current_engagement_score'] = st.number_input(
                    "Current Engagement Score (1-10)", 
                    min_value=1.0, 
                    max_value=10.0,
                    value=params['current_engagement_score'],
                    step=0.1,
                    key=f"engagement_{initiative_key}"
                )
            
            with col2:
                params['engagement_improvement'] = st.slider(
                    "Engagement Score Improvement", 
                    0.0, 3.0, 
                    params['engagement_improvement'],
                    step=0.1,
                    key=f"engagement_improve_{initiative_key}"
                )
                params['retention_improvement'] = st.slider(
                    "Retention Improvement (%)", 
                    0, 50, 
                    params['retention_improvement'],
                    key=f"retention_improve_{initiative_key}"
                )
            
            results = calculate_engagement_roi(params)
            
        elif initiative_key == 'talent_development':
            with col1:
                params['development_participants'] = st.number_input(
                    "Development Participants", 
                    min_value=1, 
                    value=params['development_participants'],
                    key=f"dev_participants_{initiative_key}"
                )
                params['internal_mobility_rate'] = st.number_input(
                    "Current Internal Mobility Rate (%)", 
                    min_value=0, 
                    max_value=100,
                    value=params['internal_mobility_rate'],
                    key=f"mobility_{initiative_key}"
                )
            
            with col2:
                params['mobility_improvement'] = st.slider(
                    "Mobility Improvement (%)", 
                    0, 50, 
                    params['mobility_improvement'],
                    key=f"mobility_improve_{initiative_key}"
                )
                params['succession_improvement'] = st.slider(
                    "Succession Planning Improvement (%)", 
                    0, 50, 
                    params['succession_improvement'],
                    key=f"succession_{initiative_key}"
                )
            
            results = calculate_development_roi(params)
    
    # Display results
    st.subheader("📈 Results")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "ROI",
            f"{results['roi']:.0f}%",
            delta=get_roi_status(results['roi'])
        )
    
    with col2:
        investment_key = 'total_investment' if 'total_investment' in results else 'total_costs'
        st.metric(
            "Total Investment",
            format_currency(results[investment_key])
        )
    
    with col3:
        benefits_key = 'annual_benefits' if 'annual_benefits' in results else 'annual_savings'
        st.metric(
            "Annual Benefits",
            format_currency(results[benefits_key])
        )
    
    with col4:
        if 'payback_months' in results:
            st.metric(
                "Payback Period",
                f"{results['payback_months']:.1f} months"
            )
        else:
            net_benefit = results[benefits_key] - results[investment_key]
            st.metric(
                "Net Annual Benefit",
                format_currency(net_benefit)
            )
    
    # Benefits breakdown chart
    if 'benefit_breakdown' in results:
        breakdown = results['benefit_breakdown']
        
        fig = px.bar(
            x=list(breakdown.keys()),
            y=list(breakdown.values()),
            title=f"{template['name']} - Benefits Breakdown",
            color=list(breakdown.values()),
            color_continuous_scale="Viridis"
        )
        fig.update_layout(
            xaxis_title="Benefit Category",
            yaxis_title="Annual Value ($)",
            showlegend=False
        )
        st.plotly_chart(fig, use_container_width=True)

def display_overall_summary():
    """Display summary across all selected initiatives"""
    st.subheader("🎯 Overall Portfolio Summary")
    
    total_investment = 0
    total_benefits = 0
    initiative_results = []
    
    for initiative_key in st.session_state.selected_initiatives:
        params = st.session_state.params[initiative_key]
        
        if initiative_key in ['leadership_development', 'executive_coaching']:
            results = calculate_leadership_roi(params)
            investment = results['total_costs']
            benefits = results['annual_benefits']
        elif initiative_key == 'recruiting_optimization':
            results = calculate_recruiting_roi(params)
            investment = results['total_investment']
            benefits = results['annual_savings']
        elif initiative_key == 'onboarding_excellence':
            results = calculate_onboarding_roi(params)
            investment = results['total_investment']
            benefits = results['annual_benefits']
        elif initiative_key == 'engagement_retention':
            results = calculate_engagement_roi(params)
            investment = results['total_investment']
            benefits = results['annual_benefits']
        elif initiative_key == 'talent_development':
            results = calculate_development_roi(params)
            investment = results['total_investment']
            benefits = results['annual_benefits']
        
        total_investment += investment
        total_benefits += benefits
        
        initiative_results.append({
            'Initiative': INITIATIVE_TEMPLATES[initiative_key]['name'],
            'Investment': investment,
            'Annual Benefits': benefits,
            'ROI (%)': results['roi']
        })
    
    overall_roi = ((total_benefits - total_investment) / total_investment * 100) if total_investment > 0 else 0
    
    # Overall metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Portfolio ROI", f"{overall_roi:.0f}%", delta=get_roi_status(overall_roi))
    
    with col2:
        st.metric("Total Investment", format_currency(total_investment))
    
    with col3:
        st.metric("Total Annual Benefits", format_currency(total_benefits))
    
    with col4:
        st.metric("Net Annual Benefit", format_currency(total_benefits - total_investment))
    
    # Initiative comparison
    df = pd.DataFrame(initiative_results)
    df = df.sort_values('ROI (%)', ascending=False)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Initiative Comparison")
        st.dataframe(df, use_container_width=True)
    
    with col2:
        fig = px.bar(
            df,
            x='Initiative',
            y='ROI (%)',
            title="ROI by Initiative",
            color='ROI (%)',
            color_continuous_scale="RdYlGn"
        )
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    # Export functionality
    st.divider()
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📋 Export Summary Report", type="primary"):
            report = create_summary_report(initiative_results, overall_roi, total_investment, total_benefits)
            st.download_button(
                label="📥 Download Report",
                data=report,
                file_name=f"hr_roi_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                mime="text/plain"
            )
    
    with col2:
        if st.button("📊 Export Data (JSON)", type="secondary"):
            export_data = {
                'initiatives': initiative_results,
                'summary': {
                    'total_investment': total_investment,
                    'total_benefits': total_benefits,
                    'overall_roi': overall_roi
                },
                'timestamp': datetime.now().isoformat()
            }
            st.download_button(
                label="📥 Download JSON",
                data=json.dumps(export_data, indent=2, default=str),
                file_name=f"hr_roi_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json"
            )

def create_summary_report(initiative_results, overall_roi, total_investment, total_benefits):
    """Create a summary report"""
    report = f"""
HR ROI CALCULATOR - SUMMARY REPORT
Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}

EXECUTIVE SUMMARY
================
Portfolio ROI: {overall_roi:.0f}%
Total Investment: {format_currency(total_investment)}
Total Annual Benefits: {format_currency(total_benefits)}
Net Annual Benefit: {format_currency(total_benefits - total_investment)}

Status: {get_roi_status(overall_roi)}

INITIATIVE BREAKDOWN
===================
"""
    
    for initiative in initiative_results:
        report += f"""
{initiative['Initiative']}:
  Investment: {format_currency(initiative['Investment'])}
  Annual Benefits: {format_currency(initiative['Annual Benefits'])}
  ROI: {initiative['ROI (%)']:.0f}%
"""
    
    report += f"""

RECOMMENDATIONS
===============
{"✅ Strong portfolio - proceed with implementation" if overall_roi >= 200 else "⚠️ Review high-performing initiatives for priority implementation" if overall_roi >= 100 else "❌ Reassess assumptions and focus on highest ROI initiatives"}

Generated by HR ROI Calculator
"""
    
    return report

if __name__ == "__main__":
    main()
