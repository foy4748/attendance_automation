"""BINARY SEARCH ALGORITHM"""

def bs(arr, item):
    if len(arr) < 1:
        return False

    mid = len(arr)//2
    if arr[mid] == item:
        return True

    if item in arr[:mid]:
        bs(arr[:mid],item)
    elif item in arr[mid:]:
        bs(arr[mid:],item)
    else:
        return False

    return True
