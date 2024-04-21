from expyvalidations.exceptions import ValidateException


def validate_cpf(cpf: str) -> str:
    # Remove any non-digit characters from the input
    cpf = "".join(filter(str.isdigit, cpf))

    # Check if the CPF has 11 digits
    if len(cpf) != 11:
        raise ValidateException("Invalid CPF")

    # Calculate the first verification digit
    sum = 0
    for i in range(9):
        sum += int(cpf[i]) * (10 - i)
    digit1 = 11 - (sum % 11)
    if digit1 > 9:
        digit1 = 0

    # Calculate the second verification digit
    sum = 0
    for i in range(10):
        sum += int(cpf[i]) * (11 - i)
    digit2 = 11 - (sum % 11)
    if digit2 > 9:
        digit2 = 0

    # Check if the verification digits match the input CPF
    if int(cpf[9]) == digit1 and int(cpf[10]) == digit2:
        return cpf
    else:
        raise ValidateException("Invalid CPF")
