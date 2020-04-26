from datetime import date


def set_risk_adjusts():
    """creates dictionary object containing risk adjustments 1-6 are base risks as per protocol, age adjust_per decade adds risk per decade from 50 to 80, fem sex removes risk from females over 60. prior ix scales risk if ct in last year or colonoscopy in last 3 years, time_inc adds risk weighting for each week on waiting list"""
    risk_adjust = {
        1: 20,
        2: 16,
        3: 12,
        4: 8,
        5: 4,
        6: 0,
        "age_adjust_per_decade": 1,
        "fem_sex": -1,
        "prior_ix": 0.75,
        "time_inc": 1
    }
    return risk_adjust


def make_patient():
    """Currently creates dummy test patient dictionary object- should be replaced by api"""
    patient = {
        "hosp_no": 1234,
        "forename": "Joe",
        "surname": "Bloggs",
        "dob": date(1950, 7, 7),
        "female": False,
        "palp_mass": False,
        "ct_mass": False,
        "bleeding": True,
        "loose_stools": True,
        "fit_done": False,
        "fit_value": 0,
        "prior_colonoscopy": False,
        "prior_CT": False,
        "date_listed": date(2020, 3, 25),
        "ct_pending":True,
        "risk": 0
    }
    return patient


def symptom_risk(patient, risk_adjust):
    """calculates and returns baseline risk from symptoms and FIT result using risk weightings specified in risk_adjust dictionary """
    if patient["palp_mass"] or patient["ct_mass"]:
        return risk_adjust[1]
    elif (patient["loose_stools"] and patient["bleeding"]) or (patient["fit_value"] > 100):
        return risk_adjust[2]
    elif patient["bleeding"]:
        return risk_adjust[3]
    elif patient["loose_stools"] and patient["fit_value"] > 10:
        return risk_adjust[4]
    elif patient["fit_value"] > 10:
        return risk_adjust[5]
    else:
        return risk_adjust[6]


def age_sex_risk_adjust(patient, risk_adjust):
    """adjusts risk up for age over 50 and down for female according to risk adjust settings - returns adjustment factor to be applied to patient risk outside the function """
    this_day = date.today()
    patient_dob = (patient["dob"])
    age_object = this_day - patient_dob
    age = age_object.days / 365
    if age < 50:
        unisex_risk = 0
    elif age > 80:
        unisex_risk = 4 * risk_adjust["age_adjust_per_decade"]
    else:
        unisex_risk = int((age - 50) / 10 * risk_adjust["age_adjust_per_decade"])
    if age > 60 and patient["female"]:
        return unisex_risk - risk_adjust["fem_sex"]
    else:
        return unisex_risk


def prior_investigation_risk_adjust(patient, risk_adjust):
    """scales risk down if previous clear ct in last year or clear colonoscopy in last 3 years"""
    if patient["prior_CT"] or patient["prior_colonoscopy"]:
        return patient["risk"] * risk_adjust["prior_ix"]
    else:
        return patient["risk"]


def waiting_time_at_entry_adjust(patient, risk_adjust):
    """calculates and returns increment for weeks on waiting list"""
    date_listed = patient["date_listed"]
    date_entered = date.today()
    wait_object = date_entered - date_listed
    weeks_waiting_at_entry = int(wait_object.days / 7)
    return weeks_waiting_at_entry * risk_adjust["time_inc"]

def write_to_Db(patient):
    """writes inputted data to database"""
    pass
    #tables for 1. demographics, 2. symptoms and prior ix, 3 FIT status
    #?do in SQLLITE or ?via SQLALCHEMY

def find_needing_fit():
    """runs dB query to find patients in DB FIT table with no FIT test and no documented bleeding and outputs list"""
    pass

def enter_fit_result():
    """Allows entry of FIT test result and re-runs risk assessment"""
    pass



def risk_adjustment_on_entry(patient, risk_adjust):
    """main control function for adjusting patient risk on entry"""
    risk1 = symptom_risk(patient, risk_adjust)
    patient["risk"] = risk1
    print("stage 1 risk(symptoms and FIT): ", patient["risk"])
    risk2 = age_sex_risk_adjust(patient, risk_adjust)
    patient["risk"] += risk2
    print("stage 2 risk(age & sex adjusted):", patient["risk"])
    risk3 = prior_investigation_risk_adjust(patient, risk_adjust)
    patient["risk"] = risk3
    print("stage 3 risk(prior Ix adjusted):", patient["risk"])
    wait_time_adjust = waiting_time_at_entry_adjust(patient, risk_adjust)
    patient["risk"] += wait_time_adjust
    print("stage 4 risk(wait time adjusted):", patient["risk"])

    return patient


# main
risk_adjust = set_risk_adjusts()
patient = make_patient()
risk_adjusted_patient = risk_adjustment_on_entry(patient, risk_adjust)

#output to console for demo purpose
if patient["fit_done"]:
    provisional = "Risk includes FIT test"
elif patient["bleeding"]:
    provisional = "FIT test not required"
else:
    provisional = "Provisional Risk Pending FIT test"
if patient["ct_pending"]: 
    provisional = provisional +"      (CT currently arranged - result pending)"
print()
print(patient["forename"], patient["surname"], patient["hosp_no"], "   Risk ", patient["risk"], "     ", provisional)
