def canon(n):
    can = []
    for d in range(2, n + 1):
        k = 0
        while n % d == 0:
            n //= d
            k += 1
        if k > 0:
            can.append((d, k))
    return can


def smallest_common(a, b):
    a_canon = canon(a)
    b_canon = canon(b)
    scm = a * b
    for ca in a_canon:
        for cb in b_canon:
            if ca[0] == cb[0]:
                scm //= ca[0] ** min(ca[1], cb[1])
    return scm


def lcm(a, b, c):
    return smallest_common(smallest_common(a, b), c)


q = 0
for a in range(1, 390):
    for b in range(1, 390):
        for c in range(1, 390):
            if c > b:
                if b > a:
                    lc = lcm(a, b, c)
                    if c < lc:
                        if lc < a + b + c:
                            print(a, b, c)
                            q += 1
print(q)