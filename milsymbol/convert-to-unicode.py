import sys
f = open(sys.argv[1], mode='r', encoding='utf-8')
for c in f.read():
  if ord(c) <= 0x7F:
    print(c, end='')
  else:
    print(f'\\u{ord(c):04X}', end='')