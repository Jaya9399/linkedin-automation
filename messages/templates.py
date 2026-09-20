from datetime import datetime


def now_string():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def referral_message(person_name, company, job_title, job_id="", job_url="", matched_skills=None):
    matched_skills = matched_skills or []
    skill_text = ""
    if matched_skills:
        skill_text = "\n\nMy experience with " + ", ".join(matched_skills[:6]) + " aligns well with this role."

    job_line = f"\n\nJob: {job_url}" if job_url else ""
    return f"""Hi {person_name},

I came across the {job_title} opportunity at {company} (Job ID: {job_id}) and wanted to reach out regarding a possible referral.

I have 1.5+ years of professional experience as a Software Engineer, with hands-on experience in areas relevant to this position.{skill_text}

I have already shared my resume along with the relevant job details for your reference.{job_line}

If you feel my profile would be a good fit, I would really appreciate your help with a referral.

Thank you for your time and consideration!"""


def connection_message(person_name, company, job_title):
    return f"""Hi {person_name},

I came across your profile while looking into opportunities at {company}. I'm currently exploring Software Engineering opportunities, particularly roles such as {job_title}, and would be glad to connect.

Thanks!"""


def referral_followup_message(person_name, company, job_title, job_id="", job_url=""):
    job_line = f"\n\nJob: {job_url}" if job_url else ""
    return f"""Hi {person_name},

Thank you for connecting.

I wanted to follow up regarding the {job_title} opportunity at {company} (Job ID: {job_id}).{job_line}

If you are comfortable referring me for this position, I would really appreciate your help.

Thank you again for your time!"""


def connection_accepted_message(person_name, company, job_title, job_id="", job_url=""):
    job_line = f"\n\nJob: {job_url}" if job_url else ""
    return f"""Hi {person_name},

Thank you for connecting.

I had reached out regarding the {job_title} opportunity at {company} (Job ID: {job_id}).{job_line}

If you are familiar with the team or hiring process and feel my profile could be a fit, I would really appreciate any guidance or referral support.

Thank you!"""


def prepare_referral_already_received_message(person_name, company, job_title, job_id=""):
    return f"""Hi {person_name},

Thank you so much for being willing to help with the referral for the {job_title} role at {company}.

I wanted to let you know that I have already received a referral for this position from someone else, so I won't need to trouble you further.

I really appreciate your willingness to help!

Thank you again."""


def thank_you_referral_message(person_name, company, job_title):
    return f"""Hi {person_name},

Thank you so much for helping me with the referral for the {job_title} role at {company}.

I really appreciate your time and support.

Thank you again!"""


def alternative_process_acknowledgement(person_name, company):
    return f"""Hi {person_name},

Thank you for sharing the information regarding the application process at {company}.

I really appreciate your guidance. I'll follow the process you suggested.

Thank you again!"""


def declined_acknowledgement(person_name):
    return f"""Hi {person_name},

Thank you for getting back to me. I really appreciate you taking the time to respond.

Thanks again, and I hope you have a great day!"""


def ambiguous_response_message(person_name):
    return f"""Hi {person_name},

Thank you for getting back to me. I really appreciate your response.

Thanks again!"""
