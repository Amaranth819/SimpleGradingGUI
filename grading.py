import tkinter
import netids
import json
import os
import xlwt


class Question(object):
    def __init__(self, name : str, points : float) -> None:
        '''
            name: str
            points: float
        '''
        self.name = name
        self.points = points
        self.grade_details = {}
    

    def add_grade_details(self, choice, lose_points):
        idx = len(self.grade_details.keys())
        self.grade_details[str(idx)] = (choice, lose_points)


    def get_choice_details(self):
        return {idx : choice for idx, (choice, _) in self.grade_details.items()}


    def get_lose_points(self):
        return {idx : lose_points for idx, (_, lose_points) in self.grade_details.items()}


class Rubrics(object):
    def __init__(self) -> None:
        self.questions = {}


    def add_question(self, q : Question):
        self.questions[q.name] = q


    def all_question_names(self):
        return self.questions.keys()


    def init_feedback(self):
        choice_data, extra_comment = {}, {}
        for qn in self.all_question_names():
            choice_data[qn] = {idx : False for idx in self.questions[qn].grade_details.keys()}
            extra_comment[qn] = ''
        return choice_data, extra_comment



class FeedbackSummary(object):
    def __init__(self, student_netids, rubrics : Rubrics) -> None:
        self.choices = {}
        self.extra_comments = {}
        self.rubrics = rubrics
        self.student_netids = student_netids

        for nid in student_netids:
            self.choices[nid], self.extra_comments[nid] = rubrics.init_feedback()


    def calculate_points(self):
        def _helper(dict1, dict2):
            # Both dictionaries have same keys.
            res = 0
            for key in dict1.keys():
                res += dict1[key] * dict2[key]
            return res
            
        all_points = {}
        for netid in self.student_netids:
            curr_points = {}
            for qn in self.rubrics.all_question_names():
                total_points = self.rubrics.questions[qn].points
                choices_points = self.rubrics.questions[qn].get_lose_points()
                choices_made = self.choices[netid][qn]
                curr_points[qn] = total_points + _helper(choices_points, choices_made)
            all_points[netid] = curr_points
        return all_points


    def generate_comments(self):
        def _helper(dict1, dict2):
            # Both dictionaries have same keys.
            res = []
            for key in dict1.keys():
                if dict2[key] == True:
                    res.append(dict1[key])
            return ','.join(res)

        comments = {}
        for netid in self.student_netids:
            current_comment = {}
            for qn in self.rubrics.all_question_names():
                choices_details = self.rubrics.questions[qn].get_choice_details()
                choices_made = self.choices[netid][qn]
                current_comment[qn] = _helper(choices_details, choices_made)
                if self.extra_comments[netid][qn] != '':
                    if current_comment[qn] == '':
                        current_comment[qn] = self.extra_comments[netid][qn]
                    else:
                        current_comment[qn] += ':' + self.extra_comments[netid][qn]
            comments[netid] = current_comment
        return comments


    def save_as_json(self, json_path):
        js = json.dumps({'choices' : self.choices, 'extra_comments' : self.extra_comments})
        with open(json_path, 'w') as j:
            j.write(js)


    def load_json(self, json_path):
        with open(json_path) as j:
            f = json.load(j)
            self.choices = f['choices']
            self.extra_comments = f['extra_comments']


    def export_grade_to_excel(self, excel_path):
        points, comments = self.calculate_points(), self.generate_comments()

        workbook = xlwt.Workbook(encoding = 'utf-8')
        sheet = workbook.add_sheet('feedbacks', cell_overwrite_ok = True)

        sheet.write(0, 0, 'NetID')
        qnames = self.rubrics.all_question_names()
        col_idx = 0
        for qname in qnames:
            col_idx += 1
            sheet.write(0, col_idx, qname)
            col_idx += 1
            sheet.write(0, col_idx, f'{qname} Feedback')

        for r, netid in enumerate(self.student_netids):
            sheet.write(r+1, 0, netid)
            col_idx = 0
            for qname in qnames:
                col_idx += 1
                sheet.write(r+1, col_idx, points[netid][qname])
                col_idx += 1
                sheet.write(r+1, col_idx, f'{qname}:' + comments[netid][qname])

        workbook.save(excel_path)


'''
    Create UI
'''
def create_ui(root, rubrics : Rubrics, n_col = 2):
    q_vars = {}
    ec_vars = {}

    for e, qn in enumerate(rubrics.all_question_names()):
        q_var = {}
        labelframe = tkinter.LabelFrame(root)
        labelframe.grid(row = e // n_col + 1, column = e % n_col)

        q_label = tkinter.Label(labelframe, text = qn)
        q_label.pack()

        for idx, (detail, _) in rubrics.questions[qn].grade_details.items():
            val = tkinter.IntVar()
            cb = tkinter.Checkbutton(
                labelframe, 
                text = f'{detail}', 
                variable = val, 
                onvalue = True, 
                offvalue = False
            )
            cb.pack()
            q_var[str(idx)] = val
        
        ec_label = tkinter.Label(labelframe, text = 'Extra comment:')
        ec_label.pack()
        
        ec_var = tkinter.StringVar()
        input_txt = tkinter.Entry(labelframe, textvariable = ec_var)
        input_txt.pack()
        ec_vars[qn] = ec_var

        q_vars[qn] = q_var

    return q_vars, ec_vars



def vars_to_vals(q_vars):
    q_vals = {}

    for k in q_vars.keys():
        q_vals[k] = {}
        for idx in q_vars[k].keys():
            q_vals[k][idx] = q_vars[k][idx].get()

    return q_vals


def ecvars_to_strs(ec_vars):
    ec_strs = {}

    for k in ec_vars.keys():
        ec_strs[k] = ec_vars[k].get()

    return ec_strs



'''
    Create rubrics
'''
def hw1_rubrics():
    q1 = Question('Q3(a)', 2)
    q1.add_grade_details('Correct', 0)
    q1.add_grade_details('Partially correct (-0.5)', -0.5)
    q1.add_grade_details('No answer (-2)', -2)
    q2 = Question('Q3(b)', 3)
    q2.add_grade_details('Correct', 0)
    q2.add_grade_details('Partially correct (-1)', -1)
    q2.add_grade_details('No answer (-3)', -3)

    rubrics = Rubrics()
    rubrics.add_question(q1)
    rubrics.add_question(q2)

    return rubrics



if __name__ == '__main__':
    rubrics = hw1_rubrics()

    student_netids = netids.netids
    summary = FeedbackSummary(student_netids, rubrics)

    json_path = 'feedback.json'
    if os.path.exists(json_path):
        summary.load_json(json_path)

    root = tkinter.Tk()

    # Buttons for choices
    def set_values(q_vars, vals):
        for q in q_vars.keys():
            for k in q_vars[q].keys():
                q_vars[q][k].set(vals[q][k])


    def set_ecs(ec_vars, extra_comments):
        for q in ec_vars.keys():
            ec_vars[q].set(extra_comments[q])


    def load_student(netid, summary : FeedbackSummary, q_vars, ec_vars):
        netid_label.config(text = netid)
        set_values(q_vars, summary.choices[netid])
        set_ecs(ec_vars, summary.extra_comments[netid])


    def prev_student(summary, q_vars, ec_vars):
        global curr_student_idx, student_netids
        save_feedbacks(summary, q_vars, ec_vars)
        if curr_student_idx > 0:
            curr_student_idx -= 1
        load_student(student_netids[curr_student_idx], summary, q_vars, ec_vars)


    def next_student(summary, q_vars, ec_vars):
        global curr_student_idx, student_netids
        save_feedbacks(summary, q_vars, ec_vars)
        if curr_student_idx < len(student_netids) - 1:
            curr_student_idx += 1
        load_student(student_netids[curr_student_idx], summary, q_vars, ec_vars)


    def save_feedbacks(summary, q_vars, ec_vars):
        global curr_student_idx, student_netids
        summary.choices[student_netids[curr_student_idx]] = vars_to_vals(q_vars)
        summary.extra_comments[student_netids[curr_student_idx]] = ecvars_to_strs(ec_vars)

    curr_student_idx = 0
    q_vars, ec_vars = create_ui(root, rubrics)

    info_labelframe = tkinter.LabelFrame(root, width = root.winfo_width())
    info_labelframe.grid(row = 0, columnspan = 2)
    prev_button = tkinter.Button(info_labelframe, text = 'Prev', command = lambda: prev_student(summary, q_vars, ec_vars))
    prev_button.pack(side = tkinter.LEFT, fill = tkinter.BOTH, expand = True)

    netid_label = tkinter.Label(info_labelframe, text = student_netids[curr_student_idx])
    netid_label.pack(side = tkinter.LEFT, fill = tkinter.BOTH, expand = True)

    next_button = tkinter.Button(info_labelframe, text = 'Next', command = lambda: next_student(summary, q_vars, ec_vars))
    next_button.pack(side = tkinter.LEFT, fill = tkinter.BOTH, expand = True)
    
    load_student(student_netids[curr_student_idx], summary, q_vars, ec_vars)

    root.mainloop()

    summary.save_as_json(json_path)
    summary.export_grade_to_excel('feedback.xls')
