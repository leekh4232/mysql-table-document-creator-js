# MySQL 테이블 명세서 생성 프로그램


> 이 프로그램은 메가스터디IT아카데미(신촌)에서 진행되는 이광호 강사의 수업에 참여하는 훈련생들의 포트폴리오 작업에 도움을 주기 위해 작성되었습니다.

MySQL의 테이블 명세서를 생성하는 Node 프로그램입니다.

## 사용 방법

### 1) 설치

우선 Node.js가 없다면 아래 사이트에서 내려받아 설치합니다.

> https://nodejs.org/ko

터미널에서 아래의 명령을 수행합니다.

```shell
$ npm install -g https://github.com/leekh4232/mysql-table-document-creator-js.git
```

### 3) 명령프롬프트 상에서 실행

이후 이 프로그램을 운영체제 안에서 명령어로 사용할 수 있습니다.

```shell
$ tabledoc -h 호스트주소 -d DB이름 -u 사용자계정 -p 비밀번호 -output 생성될폴더경로 -port 포트번호
```

| 파라미터 | 설명 | 기본값 |
|---|---|---|
| -h | 호스트주소 | 127.0.0.1 |
| -d | DB이름 | test |
| -u | 사용자계정 | root |
| -p | 비밀번호 | 123qwe!@# |
| -output | 생성될폴더경로 | 현재 작업 디렉토리 |
| -port | 포트번호 | 3306 |

#### 사용 예시

데이터베이스 이름은 `myschool`이고 `root`계정의 비밀번호는 `1234`이며 데이터베이스 포트번호는 `9090`인 경우

그 밖의 나머지는 기본값 사용

```shell
$ tabledoc -d myschool -u root -p 1234 -port 9090
```

### 4) 생성된 엑셀 파일 확인

![res/result.png](res/result.png)

### 5) 프로그램 삭제시

```shell
npm uninstall -g tabledoc
```
