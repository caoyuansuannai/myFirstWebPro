import www.orm
import asyncio

from www import orm
from www.models import User, Blog, Comment


async def test(loop):
    await orm.create_pool(loop=loop, user='root', password='password', db='awesome')

    u = User(name='Test', email='tes1@qq.com', passwd='1234567890', image='about:blank')

    print(u)
    await u.save()
    ## 网友指出添加到数据库后需要关闭连接池，否则会报错 RuntimeError: Event loop is closed
    orm.__pool.close()
    await orm.__pool.wait_closed()


if __name__ == '__main__':
    loop = asyncio.get_event_loop()
    loop.run_until_complete(test(loop))
    loop.close()
