def bubble_sort(numbers):
    for i in range(len(numbers)):
        for j in range(len(numbers) - 1):
            if numbers[j] > numbers[j + 1]:
                numbers[j], numbers[j + 1] = numbers[j + 1], numbers[j]
    return numbers

numbers = [4, 3, 12, 1, 5, 17, 12 ,2, 6, 9, 10]

print(bubble_sort(numbers))