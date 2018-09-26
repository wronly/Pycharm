# coding=utf-8
import time, os, sched

# 第一个参数确定任务的时间，返回从某个特定的时间到现在经历的秒数
# 第二个参数以某种人为的方式衡量时间
schedule = sched.scheduler(time.time, time.sleep)


def perform_command(cmd, inc,k):
    if k==3:
        print(k)
        k=k+1
        exit(0)
    else:
        print(k)
        k = k + 1
        # 安排inc秒后再次运行自己，即周期运行
        schedule.enter(inc, 0, perform_command, (cmd, inc, k))
        os.system(cmd)


i=1
def timming_exe(cmd, inc,j):
    print(j)
    if j==3:
        exit(0)
    else:
        # enter用来安排某事件的发生时间，从现在起第n秒开始启动
        k = j + 1
        schedule.enter(inc, 0, perform_command, (cmd, inc,1))
        # 持续运行，直到计划时间队列变成空为止
        schedule.run()



print("show time after 10 seconds:")

timming_exe("echo %time%", 3,1)

#  定时器开始时间   如何结束
#