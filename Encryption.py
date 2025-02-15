# Functions
def Encrypt(target_string, key_1, key_2, key_3):
    encrypted_string = ""
    alphas = "ABCDEFGHIJKLNMOPQRSTUVWXYZabcdefghijklnmopqrstuvwxyz0123456789"
    shift_position = 0
    previous_position_culminative = 0
    for n in range(len(target_string)):
        for i in range(len(alphas)):
            if target_string[n] == alphas[i]:
                shift_position = (i + key_1 * previous_position_culminative + key_2 * n + key_3) % len(alphas)
                previous_position_culminative += shift_position
                break
        encrypted_string += alphas[shift_position]
    return encrypted_string

def Decrypt(target_string, key_1, key_2, key_3):
    decrypted_string = ""
    alphas = "ABCDEFGHIJKLNMOPQRSTUVWXYZabcdefghijklnmopqrstuvwxyz0123456789"
    shift_position = 0
    previous_position_culminative = 0
    for n in range(len(target_string)):
        for i in range(len(alphas)):
            shift_position = (i + key_1 * previous_position_culminative + key_2 * n + key_3) % len(alphas)
            if target_string[n] == alphas[shift_position]:    
                previous_position_culminative += shift_position
                break
        decrypted_string += alphas[i]
    return decrypted_string

# Test
key_1 = 7
key_2 = 6
key_3 = 3
target_string = "abcerwtewrterwtewrterw"
print("Original String: " + target_string)
encrypted_string = Encrypt(target_string, key_1, key_2, key_3)
print("Encrypted String: " + encrypted_string)
decrypted_string = Decrypt(encrypted_string, key_1, key_2, key_3)
print("Decrypted_String: " + decrypted_string)
