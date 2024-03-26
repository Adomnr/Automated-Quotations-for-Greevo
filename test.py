def round_up_to_nearest_thousand(number):
    if number % 1000 == 0:
        return number
    else:
        return ((number // 1000) + 1) * 1000

# Test the function
number = 1276890
rounded_number = round_up_to_nearest_thousand(number)
print(f"The nearest thousand of {number} is {rounded_number}")
