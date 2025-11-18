class MCO:
    def __init__(self, data, treeview):
        self.data = data
        self.criteria_num = len(data.columns) - 1
        self.treeview = treeview

        self.binary_relations = []

    def agreement_compromise(self):
        self.agreement = 'A'
        self.compromise = 'C'

        self.ac_column = 'A|C'
        if self.ac_column not in self.data.columns:
            self.data[self.ac_column] = None

        for i, row in enumerate(self.binary_relations):
            self.data.at[i, self.ac_column] = self.compromise
            for j in range(len(row)):
                if j != i:
                    betters, worses = row[j].split(':')
                    if betters == '0':
                        self.data.at[i, self.ac_column] = self.agreement
                        break
            
    def normalize(self, directions, gc_dir):
        self.agreement_compromise()
        self.normalized_postfix = '_normalized'
        for column in self.data.columns[1:self.criteria_num + 1]:
            target_column = column + self.normalized_postfix
            if target_column not in self.data.columns:
                self.data[target_column] = None
            
            for index, value in enumerate(self.data[column]):
                if self.data.at[index, self.ac_column] == self.compromise:
                    min = self.data[column].min()
                    max = self.data[column].max()
                    normilized_value = round((value - min) / (max - min), 2)
                    if directions[column] != gc_dir:
                        normilized_value = round(1 - normilized_value, 2)

                    self.data.at[index, target_column] = normilized_value
                else:
                    self.data.at[index, target_column] = ''
        
    def main_criterion(self, weights, lower_restrictions, upper_restrictions, gc_dir):
        main_column = max(weights, key=weights.get)
        mc_column = 'MC'
        if mc_column not in self.data.columns:
            self.data[mc_column] = None

        for index, row in self.data.iterrows():
            if self.data.at[index, self.ac_column] == self.compromise:
                self.data.at[index, mc_column] = row[main_column + self.normalized_postfix]
                for column in self.data.columns[1:self.criteria_num + 1]:
                    if column != main_column:
                        value = row[column + self.normalized_postfix]
                        if value > upper_restrictions[column] or value < lower_restrictions[column]:
                            self.data.at[index, mc_column] = ''
                            break
            else:
                self.data.at[index, mc_column] = ''

        best = 0 if gc_dir == 'max' else 1
        best_index = 0
        for i, value in enumerate(self.data[mc_column]):
            if value != '':
                if (gc_dir == 'max' and value > best) or (gc_dir == 'min' and value < best):
                    best = value
                    best_index = i
        
        return best_index + 1
    
    def additive(self, weights, gc_dir):
        additive_column = 'Addictive'
        if additive_column not in self.data.columns:
            self.data[additive_column] = None
        
        for index, row in self.data.iterrows():
            sum = 0
            if self.data.at[index, self.ac_column] == self.compromise:
                for column in self.data.columns[1:self.criteria_num + 1]:
                    sum += weights[column] * row[column + self.normalized_postfix]

                self.data.at[index, additive_column] = round(sum, 2)
            else:
                self.data.at[index, additive_column] = ''

        best = 0 if gc_dir == 'max' else 1
        best_index = 0
        for i, value in enumerate(self.data[additive_column]):
            if value != '':
                if (gc_dir == 'max' and value > best) or (gc_dir == 'min' and value < best):
                    best = value
                    best_index = i
        
        return best_index + 1

    def do_binary_relations(self, directions):
        self.binary_relations = []
        rows = []
        for index, row in self.data.iterrows():
            row_dict = {}
            for column_name in self.data.columns[1:self.criteria_num + 1]:
                row_dict[column_name] = row[column_name]

            rows.append(row_dict)

        for i, row in enumerate(rows):
            relation_row = []
            for j in range(len(rows)):
                if j != i:
                    row_to_compare = rows[j]
                    better_num = 0
                    for column in self.data.columns[1:self.criteria_num + 1]:
                        if row[column] == row_to_compare[column]:
                            better_num += 0.5
                        elif (directions[column] == 'max' and row[column] > row_to_compare[column]) or (directions[column] == 'min' and row[column] < row_to_compare[column]):
                            better_num += 1

                    if self.criteria_num - better_num == 0.5:
                        better_num += 0.5
                    elif better_num == 0.5:
                        better_num -= 0.5

                    relation_row_formatted = f'{better_num:.1f}:{self.criteria_num - better_num:.1f}'
                    if better_num % 1 == 0:
                        relation_row_formatted = f'{better_num:.0f}:{self.criteria_num - better_num:.0f}'

                    relation_row.append(relation_row_formatted)
                else:
                    relation_row.append('-')

            self.binary_relations.append(relation_row)
    
    def add_br_column(self):
        br_column = 'BR'
        if br_column not in self.data.columns:
            self.data[br_column] = None
        
        for i, row in enumerate(self.binary_relations):
            sum = 0
            for j in range(len(row)):
                if j != i:
                    betters, worses = row[j].split(':')
                    sum += float(betters)
            
            self.data.at[i, br_column] = sum

        best = 0
        best_index = 0
        for i, value in enumerate(self.data[br_column]):
            if value != '':
                if value > best:
                    best = value
                    best_index = i
        
        return best_index + 1