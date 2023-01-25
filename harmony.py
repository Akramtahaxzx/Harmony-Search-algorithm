import random

# Define the cities
cities = ['A', 'B', 'C', 'D', 'E','F']

# Define the distance matrix
distances = [
            [0, 11, 15, 16, 9, 9],
            [11, 0, 10, 15, 14, 10],
            [15, 10, 0, 8, 13, 9],
            [16, 15, 8, 0, 11, 10],
            [9, 14, 13, 11, 0, 6],
            [9, 10, 9, 10, 6, 0]]

# Define the harmony memory size (number of solutions to store in the memory)
HMS = 5

# Define the harmony memory consideration rate (probability of considering a solution from the memory)
HMCR = 0.5

# Define the pitch adjusting rate (how much to adjust the solution during pitch adjusting)
PAR = 0.1

# Define the number of iterations
iterations = 500
# Initialize the harmony memory with random solutions
harmony_memory = []
for i in range(HMS):
    random_solution = random.sample(cities, len(cities))
    harmony_memory.append(random_solution)

# Evaluate the fitness of each solution in the harmony memory
fitnesses = []
fitness_list = []
min_element_list = []
total = []
#اهنا غلط
for solution in harmony_memory:
    fitness = 0
    for i in range(len(solution) - 1):
        city1 = solution[i]
        city2 = solution[i + 1]
        city1_index = cities.index(city1)
        city2_index = cities.index(city2)
        distance = distances[city1_index][city2_index]
        fitness += distance
    fitnesses.append(fitness)

# Run the Harmony Search algorithm

for i in range(iterations):

    #to Get each column in row
    selected_cities = []
    PAR_list = []
    unselected_cities = []
    k=0
    for row in harmony_memory:
        selected_city = row[k]
        selected_cities.append(selected_city)
    k+=1
    new_solution = [None] * len(solution)

    for j in range(len(solution)):
        # Generate a new solution by considering a solution from the memory
        if random.uniform(0, 1) <= HMCR:
            selected_cities_list = list(set(selected_cities) - set(new_solution))
            if selected_cities_list:  # check if the result of the subtraction is not empty
                new_solution[j] = random.choice(selected_cities_list)

                # Adjust the new solution by pitch adjusting
                if random.uniform(0, 1) <= PAR:
                    index = selected_cities.index(random.choice(selected_cities))
                    PAR_list = harmony_memory[index][:]
                    for PAR_check in PAR_list:
                        if PAR_check not in new_solution:
                            new_solution[j] = PAR_check
                            break

            else:
                unselected_cities_list = list(set(cities) - set(new_solution))
                new_solution[j] = random.choice(unselected_cities_list)
        else:
            unselected_cities_list = list(set(cities) - set(new_solution))
            new_solution[j] = random.choice(unselected_cities_list)

    # Evaluate the fitness of the new solution
    fitness = 0
    for j in range(len(new_solution) - 1):
        city1 = new_solution[j]
        city2 = new_solution[j + 1]
        city1_index = cities.index(city1)
        city2_index = cities.index(city2)
        distance = distances[city1_index][city2_index]
        fitness += distance

    # Update the harmony memory
    if fitness < max(fitnesses):
        worst_solution_index = fitnesses.index(max(fitnesses))
        harmony_memory[worst_solution_index] = new_solution
        fitnesses[worst_solution_index] = fitness


    # Print the solutions and fitnesses at each iteration
    print('Iteration', i + 1)
    for solution, fit in zip(harmony_memory, fitnesses):
        print('Solution:', solution, 'Fitness:', fit)
        fitness_list.append(min(fitnesses))
    # To print the min element for each iteration
    min_element_list.append(min(fitnesses))
    min_element = min(fitnesses)
    print("Best solution in iteration ", i+1," (",min_element,")",)

# Find the best solution in the harmony memory
best_solution = harmony_memory[fitnesses.index(min(fitnesses))]
best_fitness = min(fitnesses)

# Print the result
print("")
print('\033[1m' + 'Result of HSA' + '\033[0m')
print('\033[1m' + 'Best solution:' + '\033[0m', best_solution)
print('\033[1m' + 'Best fitness' + '\033[0m', best_fitness)

#Put the result in Excel file
import xlsxwriter

workbook = xlsxwriter.Workbook('Result of HSA.xlsx')
worksheet = workbook.add_worksheet()

row = 0
column = 0

content1 = min_element_list

for item in content1:
    worksheet.write(row, column, item)

    row += 1
workbook.close()