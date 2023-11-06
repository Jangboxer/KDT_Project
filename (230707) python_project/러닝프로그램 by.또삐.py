from func import * #func 파일의 함수명,변수,클래스 모두 사용하겠다.

while True:
    menu()
    input1=input("원하는 번호를 선택하세요 : ")
    if input1=="1":
        save()
    elif input1=="2":
        month()
    elif input1=="3":
        year()
    elif input1=="4":
        simzone()
        pass
    elif input1=="5":
        predict()
        pass
    elif input1=="0":
        write_excel_template("러닝일지","러닝",datalist)
    elif input1=="7":
        break
    else:
        print("잘못입력하셨습니다.")
















